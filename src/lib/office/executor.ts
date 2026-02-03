/**
 * 统一代码执行器
 * 
 * 封装现有 execute_code 工具的完整逻辑，供新的领域工具调用
 */

import { 
  pendingExecutions, 
  executionResults, 
  getExecutionEventEmitter,
  getTaskEventEmitter,
  cleanupExpired
} from "./mcp-server";
import type { ExecutionResult, PendingExecution } from "./mcp-server";
import { getFixHint } from "./error-patterns";

/**
 * 代码复杂度检查结果
 */
interface ComplexityCheckResult {
  shouldWarn: boolean;
  shouldBlock: boolean;
  issues: string[];
  suggestions: string[];
  metrics: {
    lines: number;
    insertCalls: number;
    syncCalls: number;
    hasRiskyAPIs: boolean;
  };
}

/**
 * 危险 API 列表
 */
const RISKY_APIS = [
  { pattern: /body\.clear\(\)/g, name: 'body.clear()', risk: 'high', suggestion: '避免清空整个文档，在空白文档开始操作' },
  { pattern: /insertParagraph\([^)]*,\s*["']Start["']\)/g, name: 'insertParagraph(..., "Start")', risk: 'high', suggestion: '使用 "End" 顺序添加内容，避免打乱结构' },
  { pattern: /insertField\([^)]*FieldType\.toc/g, name: 'insertField(toc)', risk: 'medium', suggestion: '目录字段不稳定，考虑手动创建目录列表' },
  { pattern: /\.search\([^)]+\)\..*insert/g, name: 'search().insert*()', risk: 'medium', suggestion: '搜索定位不可靠，建议保存引用后使用 insertParagraph("After")' },
  { pattern: /shading\.backgroundPatternColor/g, name: 'shading.backgroundPatternColor', risk: 'medium', suggestion: '某些 Word 版本不支持，使用 font.highlightColor 代替' },
];

/**
 * 检查代码复杂度
 */
function checkCodeComplexity(code: string): ComplexityCheckResult {
  const issues: string[] = [];
  const suggestions: string[] = [];
  
  // 计算代码行数（忽略空行和注释）
  const meaningfulLines = code.split('\n')
    .filter(line => {
      const trimmed = line.trim();
      return trimmed.length > 0 && !trimmed.startsWith('//');
    }).length;
  
  // 计算 insert* 操作数
  const insertCalls = (code.match(/\.insert[A-Z][a-zA-Z]*\(/g) || []).length;
  
  // 计算 context.sync() 调用数
  const syncCalls = (code.match(/context\.sync\(\)/g) || []).length;
  
  // 检查危险 API
  let hasRiskyAPIs = false;
  for (const api of RISKY_APIS) {
    if (api.pattern.test(code)) {
      hasRiskyAPIs = true;
      issues.push(`⚠️ 检测到风险 API: ${api.name} (${api.risk === 'high' ? '高危' : '中危'})`);
      suggestions.push(api.suggestion);
    }
  }
  
  // 复杂度检查
  const shouldWarn = meaningfulLines > 30 || insertCalls > 5 || hasRiskyAPIs;
  const shouldBlock = meaningfulLines > 80 || insertCalls > 15;
  
  if (meaningfulLines > 30) {
    issues.push(`⚠️ 代码行数过多: ${meaningfulLines} 行（建议 ≤ 30 行）`);
    suggestions.push('将代码拆分为多个步骤，每步只处理一个章节或逻辑单元');
  }
  
  if (insertCalls > 5) {
    issues.push(`⚠️ 插入操作过多: ${insertCalls} 次（建议 ≤ 5 次）`);
    suggestions.push('减少单次执行的插入操作数量，分步执行');
  }
  
  if (syncCalls === 0) {
    issues.push('⚠️ 缺少 context.sync() 调用');
    suggestions.push('确保在操作完成后调用 await context.sync()');
  }
  
  if (syncCalls > 3) {
    issues.push(`⚠️ context.sync() 调用过多: ${syncCalls} 次（可能影响性能）`);
    suggestions.push('合并操作，减少 sync() 调用次数');
  }
  
  // 检查是否有返回验证信息
  const hasReturnValidation = /return\s*\{[\s\S]*success[\s\S]*\}/g.test(code);
  if (!hasReturnValidation) {
    issues.push('⚠️ 缺少验证返回值');
    suggestions.push('添加 return { success: true, created: "..." } 以便验证执行结果');
  }
  
  return {
    shouldWarn,
    shouldBlock,
    issues,
    suggestions,
    metrics: {
      lines: meaningfulLines,
      insertCalls,
      syncCalls,
      hasRiskyAPIs
    }
  };
}

/**
 * 生成复杂度警告消息
 */
function formatComplexityWarning(result: ComplexityCheckResult): string {
  let message = `## ⚠️ 代码复杂度警告\n\n`;
  message += `### 检测到的问题\n\n`;
  message += result.issues.map(issue => `- ${issue}`).join('\n');
  message += `\n\n### 代码指标\n\n`;
  message += `- 代码行数: ${result.metrics.lines} 行\n`;
  message += `- 插入操作: ${result.metrics.insertCalls} 次\n`;
  message += `- sync() 调用: ${result.metrics.syncCalls} 次\n`;
  message += `- 包含风险 API: ${result.metrics.hasRiskyAPIs ? '是' : '否'}\n`;
  message += `\n### 建议\n\n`;
  message += result.suggestions.map((s, i) => `${i + 1}. ${s}`).join('\n');
  message += `\n\n### 请求\n\n`;
  message += `请根据上述建议简化代码，拆分为多个步骤后重新提交。每步代码应：\n`;
  message += `- 不超过 30 行\n`;
  message += `- 不超过 5 个 insert* 操作\n`;
  message += `- 包含验证返回值 \`return { success: true, created: "..." }\`\n`;
  message += `- 只处理一个逻辑单元（如一个章节）`;
  
  return message;
}

/**
 * 统一执行 Office.js 代码
 * 
 * @param host - 目标应用 (excel/word/powerpoint)
 * @param code - 要执行的 Office.js 代码
 * @param description - 操作描述
 * @returns MCP工具响应格式
 */
export async function executeOfficeCode(
  host: 'excel' | 'word' | 'powerpoint',
  code: string,
  description?: string
) {
  const correlationId = crypto.randomUUID();
  const startTime = Date.now();
  
  console.log(`[MCP/office] Executing code in ${host}, correlationId: ${correlationId}`);
  console.log(`[MCP/office] Code length: ${code.length}`);
  if (description) {
    console.log(`[MCP/office] Description: ${description}`);
  }
  
  // 代码复杂度检查
  const complexityResult = checkCodeComplexity(code);
  
  console.log(`[MCP/office] Complexity check:`, {
    lines: complexityResult.metrics.lines,
    insertCalls: complexityResult.metrics.insertCalls,
    shouldWarn: complexityResult.shouldWarn,
    shouldBlock: complexityResult.shouldBlock,
    hasRiskyAPIs: complexityResult.metrics.hasRiskyAPIs
  });
  
  // 如果代码过于复杂，阻止执行并返回拆分建议
  if (complexityResult.shouldBlock) {
    console.warn(`[MCP/office] Code complexity too high, blocking execution`);
    return {
      content: [{
        type: "text" as const,
        text: `❌ 代码复杂度过高，已阻止执行\n\n${formatComplexityWarning(complexityResult)}`
      }]
    };
  }
  
  // 如果有警告，记录但继续执行
  if (complexityResult.shouldWarn) {
    console.warn(`[MCP/office] Code complexity warning:`, complexityResult.issues);
  }
  
  // 清理过期数据
  cleanupExpired();
  
  // 将任务放入待处理队列
  pendingExecutions.set(correlationId, {
    host: host,
    code: code,
    description: description,
    status: 'pending',
    timestamp: Date.now()
  });
  
  // 触发新任务事件（用于 SSE 推送）
  const taskEventEmitter = getTaskEventEmitter();
  if (taskEventEmitter) {
    taskEventEmitter.emit('new-task', {
      correlationId,
      host: host,
      code: code,
      description: description
    });
  }
  
  // 使用 EventEmitter 事件驱动等待结果（替代轮询，零延迟）
  const maxWait = 60000 * 5; // 60秒超时
  
  try {
    // 创建 Promise 等待事件通知
    const result = await new Promise<ExecutionResult>((resolve, reject) => {
      // 设置超时定时器
      const timeoutId = setTimeout(() => {
        // 清理监听器
        const executionEventEmitter = getExecutionEventEmitter();
        if (executionEventEmitter) {
          executionEventEmitter.removeListener(correlationId, handleResult);
        }
        reject(new Error('执行超时'));
      }, maxWait);
      
      // 结果处理函数
      const handleResult = (result: ExecutionResult) => {
        clearTimeout(timeoutId);
        resolve(result);
      };
      
      // 监听特定 correlationId 的结果事件（只触发一次）
      const executionEventEmitter = getExecutionEventEmitter();
      if (executionEventEmitter) {
        executionEventEmitter.once(correlationId, handleResult);
      } else {
        // 如果事件总线未初始化，降级到轮询模式
        console.warn('[MCP/office] EventEmitter not initialized, fallback to polling');
        const pollInterval = setInterval(() => {
          const result = executionResults.get(correlationId);
          if (result) {
            clearInterval(pollInterval);
            clearTimeout(timeoutId);
            resolve(result);
          }
        }, 100);
      }
    });
    
    // 获取到结果，清理状态
    executionResults.delete(correlationId);
    pendingExecutions.delete(correlationId);
    
    const duration = Date.now() - startTime;
    
    if (result.success) {
      console.log(`[MCP/office] Code executed successfully, correlationId: ${correlationId}, duration: ${duration}ms`);
      console.log(`[MCP/office] Result type: ${typeof result.data}, hasData: ${result.data !== undefined}`);
      
      // 构建成功消息
      let successMessage = `✅ 代码执行成功！`;
      
      // 如果有返回数据，显示
      if (result.data) {
        successMessage += `\n\n返回数据: ${JSON.stringify(result.data, null, 2)}`;
      } else {
        // 成功但无数据：提示 Agent 在读取类操作中必须 return 数据
        const desc = (description ?? '').trim();
        const isReadTask = /读取|获取|查看/.test(desc);
        if (isReadTask) {
          successMessage += `\n\n⚠️ 当前为读取类操作，但未返回数据。请确保代码在 \`context.sync()\` 之后 **return** 读取结果（例如 \`return range.values\` 或 \`return { values: range.values }\`），然后重新调用工具。`;
        } else {
          successMessage += `\n\n💡 本次执行未返回数据。若需将读取到的内容回传给 AI，请在生成的代码中在 \`context.sync()\` 之后 **return** 数据（例如 \`return range.values\` 或 \`return { values: range.values }\`）。`;
        }
      }
      
      // 如果有复杂度警告，附加提示
      if (complexityResult.shouldWarn) {
        successMessage += `\n\n---\n\n## 💡 优化建议\n\n`;
        successMessage += `虽然执行成功，但检测到以下可优化点：\n\n`;
        successMessage += complexityResult.issues.map(issue => `- ${issue}`).join('\n');
        successMessage += `\n\n下次执行类似任务时，建议：\n`;
        successMessage += complexityResult.suggestions.slice(0, 3).map((s, i) => `${i + 1}. ${s}`).join('\n');
      }
      
      return {
        content: [{
          type: "text" as const,
          text: successMessage
        }]
      };
    } else {
      // 执行失败，返回详细的错误信息让 Agent 可以修复
      console.log(`[MCP/office] Code execution failed, correlationId: ${correlationId}, duration: ${duration}ms`);
      console.log(`[MCP/office] Error type: ${result.error?.type}, message: ${result.error?.message}`);
      
      const errorType = result.error?.type || 'UnknownError';
      const fixHint = getFixHint(errorType);
      
      return {
        content: [{
          type: "text" as const,
          text: `❌ 代码执行失败

## 错误信息

- **错误类型**: ${errorType}
- **错误消息**: ${result.error?.message || '未知错误'}
${result.error?.code ? `- **错误代码**: ${result.error.code}` : ''}
${result.error?.stackTrace ? `\n**堆栈信息**:\n\`\`\`\n${result.error.stackTrace}\n\`\`\`` : ''}

## 失败的代码

\`\`\`javascript
${code}
\`\`\`

${fixHint}

## 请求

请分析上述错误，修正代码后重新调用相应的工具提交修正版本。

关键要求：
1. 分析错误类型和消息，确定根本原因
2. 参考修复建议应用相应的解决方案
3. 添加必要的错误检查（如 getItemOrNullObject、isNullObject 检查）
4. 确保使用正确的 API 参数和枚举值`
        }]
      };
    }
  } catch (error) {
    // 超时或其他错误
    const duration = Date.now() - startTime;
    console.warn(`[MCP/office] Code execution timeout, correlationId: ${correlationId}, waited: ${duration}ms`);
    pendingExecutions.delete(correlationId);
    
    return {
      content: [{
        type: "text" as const,
        text: `⏱️ 代码执行超时（60秒）

可能的原因：
- Office 应用未正确加载
- 前端与服务端连接中断
- 代码执行时间过长

建议：
- 检查 Office 应用是否正常运行
- 刷新页面后重试
- 如果代码复杂，考虑拆分为多个步骤`
      }]
    };
  }
}
