/**
 * Office 代码执行 MCP Server
 * 
 * 使用 Claude Agent SDK 的 Custom Tool 机制，
 * 让 Agent 能够通过 Tool 调用执行 Office.js 代码，
 * 并直接接收执行结果（包括错误），实现无缝的自动修复。
 * 
 * 性能优化：使用 EventEmitter 实现事件驱动的结果通知，
 * 替代原有的轮询机制，实现零延迟响应（< 1ms vs 平均 50ms）。
 */

import { EventEmitter } from "events";
import { createSdkMcpServer, tool } from "@anthropic-ai/claude-agent-sdk";
import { z } from "zod";
import { getFixHint } from "./error-patterns";

// 待执行任务队列
export interface PendingExecution {
  host: 'excel' | 'word' | 'powerpoint';
  code: string;
  description?: string;
  status: 'pending' | 'executing' | 'completed' | 'failed';
  timestamp: number;
}

// 执行结果
export interface ExecutionResult {
  success: boolean;
  data?: any;
  error?: {
    type: string;
    code?: string;
    message: string;
    stackTrace?: string;
  };
  timestamp: number;
}

// 共享状态 - 用于服务端 Tool 与前端通信
// 使用 globalThis 确保在 Next.js 热更新/按需编译时状态不会丢失
export const pendingExecutions: Map<string, PendingExecution> = 
  (globalThis as any).__mcpPendingExecutions || 
  ((globalThis as any).__mcpPendingExecutions = new Map());

export const executionResults: Map<string, ExecutionResult> = 
  (globalThis as any).__mcpExecutionResults || 
  ((globalThis as any).__mcpExecutionResults = new Map());

// 事件总线 - 用于即时通知执行结果（替代轮询）
// 使用 globalThis 确保在 Next.js 热更新/按需编译时事件总线不会丢失
const executionEventEmitter: EventEmitter = 
  (globalThis as any).__mcpExecutionEventEmitter || 
  ((globalThis as any).__mcpExecutionEventEmitter = new EventEmitter());

// 增加最大监听器数量，避免警告（默认 10）
executionEventEmitter.setMaxListeners(100);

// 任务事件总线 - 用于 SSE 推送新任务到前端
const taskEventEmitter: EventEmitter = 
  (globalThis as any).__mcpTaskEventEmitter || 
  ((globalThis as any).__mcpTaskEventEmitter = new EventEmitter());
taskEventEmitter.setMaxListeners(100);

// 清理过期数据（5分钟）
function cleanupExpired() {
  const now = Date.now();
  const expireTime = 5 * 60 * 1000;
  
  for (const [key, value] of pendingExecutions) {
    if (now - value.timestamp > expireTime) {
      pendingExecutions.delete(key);
    }
  }
  
  for (const [key, value] of executionResults) {
    if (now - value.timestamp > expireTime) {
      executionResults.delete(key);
    }
  }
}

/**
 * 创建 Office 代码执行 MCP Server
 */
export function createOfficeMcpServer() {
  return createSdkMcpServer({
    name: "office",
    version: "1.0.0",
    tools: [
      tool(
        "execute_code",
        `在 Office 应用中执行 Office.js 代码。
        
用于在 Word、Excel 或 PowerPoint 中执行操作。
代码应该是完整的、可执行的 Office.js 代码。

重要提示：
- 如果执行失败，你会收到详细的错误信息
- 请根据错误信息分析问题并重新调用此工具提交修正后的代码
- 常见错误类型包括：InvalidArgument（参数错误）、InvalidReference（引用无效）、ApiNotFound（API不可用）等`,
        {
          host: z.enum(["excel", "word", "powerpoint"]).describe("目标 Office 应用"),
          code: z.string().describe("要执行的 Office.js 代码"),
          description: z.string().optional().describe("操作描述（可选）")
        },
        async (args) => {
          const correlationId = crypto.randomUUID();
          const startTime = Date.now();
          
          console.log(`[MCP/office] Executing code in ${args.host}, correlationId: ${correlationId}`);
          console.log(`[MCP/office] Code length: ${args.code.length}`);
          if (args.description) {
            console.log(`[MCP/office] Description: ${args.description}`);
          }
          
          // 清理过期数据
          cleanupExpired();
          
          // 将任务放入待处理队列
          pendingExecutions.set(correlationId, {
            host: args.host,
            code: args.code,
            description: args.description,
            status: 'pending',
            timestamp: Date.now()
          });
          
          // 触发新任务事件（用于 SSE 推送）
          taskEventEmitter.emit('new-task', {
            correlationId,
            host: args.host,
            code: args.code,
            description: args.description
          });
          
          // 使用 EventEmitter 事件驱动等待结果（替代轮询，零延迟）
          const maxWait = 60000*5; // 60秒超时
          
          try {
            // 创建 Promise 等待事件通知
            const result = await new Promise<ExecutionResult>((resolve, reject) => {
              // 设置超时定时器
              const timeoutId = setTimeout(() => {
                // 清理监听器
                executionEventEmitter.removeListener(correlationId, handleResult);
                reject(new Error('执行超时'));
              }, maxWait);
              
              // 结果处理函数
              const handleResult = (result: ExecutionResult) => {
                clearTimeout(timeoutId);
                resolve(result);
              };
              
              // 监听特定 correlationId 的结果事件（只触发一次）
              executionEventEmitter.once(correlationId, handleResult);
            });
            
            // 获取到结果，清理状态
            executionResults.delete(correlationId);
            pendingExecutions.delete(correlationId);
            
            const duration = Date.now() - startTime;
            
            if (result.success) {
              console.log(`[MCP/office] Code executed successfully, correlationId: ${correlationId}, duration: ${duration}ms`);
              console.log(`[MCP/office] Result type: ${typeof result.data}, hasData: ${result.data !== undefined}`);
              
              return {
                content: [{
                  type: "text" as const,
                  text: `✅ 代码执行成功！${result.data ? `\n\n返回数据: ${JSON.stringify(result.data, null, 2)}` : ''}`
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
${args.code}
\`\`\`

${fixHint}

## 请求

请分析上述错误，修正代码后重新调用 \`mcp__office__execute_code\` 工具提交修正版本。

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
      )
    ]
  });
}

/**
 * 获取待执行的任务
 * 前端轮询此函数获取需要执行的代码
 */
export function getPendingExecution(): { correlationId: string; execution: PendingExecution } | null {
  for (const [correlationId, execution] of pendingExecutions) {
    if (execution.status === 'pending') {
      // 标记为执行中
      execution.status = 'executing';
      return { correlationId, execution };
    }
  }
  return null;
}

/**
 * 提交执行结果
 * 前端执行完成后调用此函数
 * 
 * 性能优化：使用 EventEmitter 立即通知等待者，实现零延迟响应
 */
export function submitExecutionResult(correlationId: string, result: ExecutionResult): boolean {
  if (!pendingExecutions.has(correlationId)) {
    console.warn(`[MCP/office] Unknown correlationId: ${correlationId}`);
    return false;
  }
  
  // 更新执行状态
  const execution = pendingExecutions.get(correlationId);
  if (execution) {
    execution.status = result.success ? 'completed' : 'failed';
  }
  
  // 触发事件，立即通知等待者（零延迟）
  executionEventEmitter.emit(correlationId, {
    ...result,
    timestamp: Date.now()
  });
  
  // 仍然保存到 Map 中，作为备份（防止事件丢失的降级方案）
  executionResults.set(correlationId, {
    ...result,
    timestamp: Date.now()
  });
  
  return true;
}

/**
 * 获取所有待执行任务（用于轮询降级方案）
 */
export function getAllPendingExecutions(): Array<{ correlationId: string; execution: PendingExecution }> {
  const result: Array<{ correlationId: string; execution: PendingExecution }> = [];
  
  for (const [correlationId, execution] of pendingExecutions) {
    if (execution.status === 'pending') {
      result.push({ correlationId, execution });
    }
  }
  
  return result;
}

/**
 * 监听新任务事件（用于 SSE 推送）
 * 
 * @param callback 任务回调函数
 * @returns 清理函数，用于移除监听器
 */
export function onNewTask(callback: (task: {
  correlationId: string;
  host: 'excel' | 'word' | 'powerpoint';
  code: string;
  description?: string;
}) => void): () => void {
  taskEventEmitter.on('new-task', callback);
  return () => taskEventEmitter.off('new-task', callback);
}
