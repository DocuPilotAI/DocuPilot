/**
 * 错误反馈构建器
 * 用于构建详细的错误反馈信息，帮助 Agent 理解错误并生成修正代码
 */

import { ExecutionError } from './code-executor';
import { getFixHint } from './error-patterns';

/**
 * 获取 Office 版本信息
 */
function getOfficeVersion(): string {
  try {
    const g = globalThis as any;
    if (g.Office?.context?.diagnostics) {
      return g.Office.context.diagnostics.version || 'unknown';
    }
  } catch {
    // 忽略错误
  }
  return 'unknown';
}

/**
 * 提取代码中使用的 API
 */
function extractUsedAPIs(code: string): string[] {
  const apis: string[] = [];
  
  // 提取 Excel API
  const excelMatches = code.match(/Excel\.\w+/g) || [];
  apis.push(...excelMatches);
  
  // 提取 Word API
  const wordMatches = code.match(/Word\.\w+/g) || [];
  apis.push(...wordMatches);
  
  // 提取 PowerPoint API
  const pptMatches = code.match(/PowerPoint\.\w+/g) || [];
  apis.push(...pptMatches);
  
  // 去重
  return Array.from(new Set(apis));
}

/**
 * 构建详细的错误反馈消息
 * 
 * @param error - 执行错误对象
 * @param code - 失败的代码
 * @param retryCount - 当前重试次数
 * @param maxRetries - 最大重试次数
 * @returns 格式化的错误反馈消息
 */
export function buildErrorFeedback(
  error: ExecutionError,
  code: string,
  retryCount: number,
  maxRetries: number
): string {
  const usedAPIs = extractUsedAPIs(code);
  const fixHint = getFixHint(error.type);
  
  let feedback = `⚠️ 代码执行失败（第 ${retryCount}/${maxRetries} 次尝试），请修正并重新生成代码：

## 错误信息

- **错误类型**: ${error.type}
- **错误消息**: ${error.message}`;

  if (error.code) {
    feedback += `\n- **错误代码**: ${error.code}`;
  }

  if (error.stackTrace) {
    feedback += `\n- **堆栈信息**: 
\`\`\`
${error.stackTrace}
\`\`\``;
  }

  feedback += `

## 失败的代码

\`\`\`javascript
${code.trim()}
\`\`\`

## Office 运行环境

- **Office 版本**: ${getOfficeVersion()}
- **平台**: ${typeof navigator !== 'undefined' ? navigator.platform : 'unknown'}
- **使用的 API**: ${usedAPIs.length > 0 ? usedAPIs.join(', ') : '无法识别'}

${fixHint}

## 修正要求

请**立即**生成修正后的完整代码（使用 HTML 注释包裹）：

1. 分析上述错误类型和消息，确定根本原因
2. 参考修复建议，应用相应的解决方案
3. 添加必要的错误检查和验证（如 getItemOrNullObject、isSetSupported）
4. 生成完整的可执行代码，不要只说明问题
5. 确保代码包含适当的 try-catch 错误处理

**重要**: 直接输出修正后的代码，格式如下：

\`\`\`
<!--OFFICE-CODE:excel
// 修正后的代码
-->
\`\`\`
`;

  return feedback;
}

/**
 * 构建最终失败消息（重试次数用尽）
 */
export function buildFinalErrorMessage(
  error: ExecutionError,
  code: string,
  retryCount: number
): string {
  return `❌ 操作失败：已尝试 ${retryCount} 次，仍然无法成功执行。

**最后的错误**: ${error.message}

**建议**:
- 检查文档是否处于正确状态（如工作表是否存在）
- 确认 Office 应用版本是否支持使用的 API
- 尝试简化操作或分步执行
- 查看详细的错误信息和代码

如需帮助，请提供更多上下文信息。`;
}
