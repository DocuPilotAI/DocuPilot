import { NextRequest } from "next/server";
import { query } from "@anthropic-ai/claude-agent-sdk";
import { getToolsForHost } from "@/lib/office/tools";
import { OfficeHostType } from "@/lib/office/host-detector";
import { createOfficeMcpServer } from "@/lib/office/mcp-server";

export async function POST(request: NextRequest) {
  // 添加请求体解析错误处理
  let prompt: string;
  let resume: string | undefined;
  let hostType: OfficeHostType;
  let mode: 'Agent' | 'Plan' = 'Agent';
  let testMode: boolean = false;
  let testCaseId: string | undefined;
  let testSessionId: string | undefined;
  let apiKey: string | undefined;
  let apiUrl: string | undefined;
  let modelName: string | undefined;

  try {
    const body = await request.json();
    prompt = body.prompt;
    resume = body.resume;
    hostType = body.hostType;
    mode = body.mode || 'Agent';
    testMode = Boolean(body.testMode);
    testCaseId = typeof body.testCaseId === "string" ? body.testCaseId : undefined;
    testSessionId = typeof body.testSessionId === "string" ? body.testSessionId : undefined;
    // 接收前端传递的 API 配置
    apiKey = body.apiKey;
    apiUrl = body.apiUrl;
    modelName = body.modelName;

    // 验证必填参数
    if (!prompt || typeof prompt !== 'string') {
      return new Response(
        JSON.stringify({ error: "缺少必填参数: prompt" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }
    if (!hostType || typeof hostType !== 'string') {
      return new Response(
        JSON.stringify({ error: "缺少必填参数: hostType" }),
        { status: 400, headers: { "Content-Type": "application/json" } }
      );
    }
  } catch (error) {
    return new Response(
      JSON.stringify({
        error: "请求体格式错误: 无法解析 JSON",
        details: error instanceof Error ? error.message : String(error)
      }),
      { status: 400, headers: { "Content-Type": "application/json" } }
    );
  }

  // 按优先级配置 API 参数
  // 1. 前端传递的非空值
  // 2. .env.local 文件配置
  // 3. 系统环境变量
  const finalApiKey = 
    (apiKey && apiKey.trim()) || 
    process.env.ANTHROPIC_API_KEY;
  
  const finalApiUrl = 
    (apiUrl && apiUrl.trim()) || 
    process.env.ANTHROPIC_BASE_URL || 
    undefined;
  
  const finalModelName = 
    (modelName && modelName.trim()) || 
    process.env.ANTHROPIC_MODEL || 
    'claude-sonnet-4-5-20250929';

  // 验证 API Key
  if (!finalApiKey) {
    return new Response(
      JSON.stringify({
        error: "ANTHROPIC_API_KEY 未配置。请在前端设置或 .env.local 文件中配置 API 密钥。"
      }),
      {
        status: 500,
        headers: { "Content-Type": "application/json" }
      }
    );
  }
  
  const encoder = new TextEncoder();

  // 根据宿主类型获取可用工具
  const allowedTools = [
    "Read", "Write", "Glob", "Grep",
    ...getToolsForHost(hostType),
    // MCP Tool：Office 代码执行
    "mcp__office__execute_code"
  ];
  
  // 创建 Office MCP Server
  const officeMcpServer = createOfficeMcpServer();
  
  // 根据模式构建系统提示
  const systemPrompt = mode === 'Plan' 
    ? `你是 DocuPilot，一个专业的 Office 任务规划助手。

## 核心原则

**重要**：你是通过 Office.js API 帮助用户操作 Office 应用（Word/Excel/PowerPoint），**不是修改项目代码**。

当前 Office 环境: ${hostType.toUpperCase()}

## Skills 位置

- Word 操作：.claude/skills/word/SKILL.md 和 TOOLS.md
- Excel 操作：.claude/skills/excel/SKILL.md 和 TOOLS.md
- PowerPoint 操作：.claude/skills/powerpoint/SKILL.md 和 TOOLS.md

## Plan 模式工作流程

在执行任何 Office 操作之前：

1. **分析需求**：理解用户想要完成的 Office 操作
2. **查阅 Skills**：阅读对应的 SKILL.md 和 TOOLS.md，了解可用的 Office.js API
3. **制定详细计划**：
   - 列出需要执行的 Office.js 操作步骤（标题、编号）
   - 说明每个步骤使用的 Office.js API
   - 预估操作结果和效果
4. **询问用户**：展示完整计划后，明确询问："是否执行此计划？（回复'是'或'执行'即可开始）"
5. **等待用户回复**：
   - 如果用户回复"是"、"执行"、"继续"、"开始"等肯定词，立即执行计划
   - 如果用户提出修改建议，调整计划后重新询问
   - 如果用户取消，礼貌告知可随时重新开始

6. **执行操作**（用户确认后）：使用 \`mcp__office__execute_code\` 工具执行代码

## 代码执行方式

使用 \`mcp__office__execute_code\` 工具执行 Office.js 代码：
- \`host\`: "${hostType}"
- \`code\`: 完整的 Office.js 代码
- \`description\`: 操作描述（可选）

**重要提示**：在用户明确同意之前，不要调用执行工具。先展示计划，等待用户确认。`
    : `你是 DocuPilot，一个智能的 Office 助手。

## 核心原则

**重要**：你是通过 Office.js API 帮助用户操作 Office 应用（Word/Excel/PowerPoint），**不是修改项目代码**。

当前 Office 环境: ${hostType.toUpperCase()}

## Skills 位置

- Word 操作：.claude/skills/word/SKILL.md 和 TOOLS.md
- Excel 操作：.claude/skills/excel/SKILL.md 和 TOOLS.md
- PowerPoint 操作：.claude/skills/powerpoint/SKILL.md 和 TOOLS.md

## ⚠️ 关键要求：使用 Tool 执行代码

对于任何 Office 操作请求，你**必须**使用 \`mcp__office__execute_code\` 工具来执行代码：

1. **参考 Skills**：读取 .claude/skills/${hostType}/TOOLS.md 中的 Office.js 代码模板
2. **调用执行工具**：使用 \`mcp__office__execute_code\` 工具，参数包括：
   - \`host\`: "${hostType}"（当前 Office 应用）
   - \`code\`: 完整的 Office.js 代码
   - \`description\`: 操作描述（可选）
3. **处理结果**：根据工具返回的结果向用户反馈

## 代码生成规则（CRITICAL）

- **必须使用 Tool**：所有 Office 操作都必须通过 \`mcp__office__execute_code\` 工具执行
- **完整性**：在一个 ${hostType.charAt(0).toUpperCase() + hostType.slice(1)}.run 块中完成所有相关操作
- **同步**：必须在所有操作完成后调用 await context.sync()
- **多段文本**：多次调用 insertParagraph，每次插入一段（不要用 \\n）
- **错误处理**：包含 try-catch 结构

## 错误处理与自动修复

当 \`mcp__office__execute_code\` 工具返回错误时，你需要：

1. **分析错误**：仔细阅读错误类型和消息
2. **修正代码**：根据错误提示修正代码
3. **重新调用**：再次调用 \`mcp__office__execute_code\` 提交修正后的代码

### 常见错误类型

1. **InvalidArgument**: 参数不正确或缺失
   - 检查参数拼写、类型、范围
   - 确保枚举值正确（如使用 Excel.ChartType.lineClusteredColumn 而不是字符串）
   - 验证单元格地址格式（如 "A1"、"B2:D4"）

2. **InvalidReference**: 引用的对象不存在
   - 使用 \`getItemOrNullObject()\` 代替 \`getItem()\`
   - 添加 \`load()\` 和 \`await context.sync()\`
   - 检查 \`isNullObject\` 确保对象有效

3. **ApiNotFound**: API 在当前环境不可用
   - 使用 \`Office.context.requirements.isSetSupported()\` 检查
   - 提供降级方案或替代实现

4. **GeneralException**: Office.js 内部错误
   - 简化操作步骤，分批执行
   - 每批后调用 \`context.sync()\`

### 错误修正示例

**错误场景**: InvalidArgument - 枚举值错误

❌ 错误代码：
\`\`\`javascript
const chart = sheet.charts.add("LineClusteredColumn", dataRange, "AutoFit");
\`\`\`

✅ 修正代码：
\`\`\`javascript
const chart = sheet.charts.add(Excel.ChartType.lineClusteredColumn, dataRange, Excel.ChartSeriesBy.columns);
\`\`\`

**错误场景**: InvalidReference - 工作表不存在

❌ 错误代码：
\`\`\`javascript
const sheet = context.workbook.worksheets.getItem("不存在的表");
\`\`\`

✅ 修正代码：
\`\`\`javascript
const sheet = context.workbook.worksheets.getItemOrNullObject("Sheet1");
sheet.load("name");
await context.sync();
if (sheet.isNullObject) {
  // 使用活动工作表作为备选
  const activeSheet = context.workbook.worksheets.getActiveWorksheet();
}
\`\`\`

现在，请根据用户的请求使用 \`mcp__office__execute_code\` 工具执行 Office.js 代码。`;
  
  // 构建完整的提示词
  const fullPrompt = `${systemPrompt}\n\n用户请求: ${prompt}`;
  
  // 创建流式输入生成器（MCP Tools 需要流式输入）
  // SDKUserMessage 需要 parent_tool_use_id 和 session_id 字段
  async function* generateMessages(): AsyncGenerator<{
    type: 'user';
    message: { role: 'user'; content: string };
    parent_tool_use_id: string | null;
    session_id: string;
  }> {
    yield {
      type: "user" as const,
      message: {
        role: "user" as const,
        content: fullPrompt
      },
      parent_tool_use_id: null,
      session_id: resume || '' // 使用 resume session 或空字符串（SDK 会自动生成）
    };
  }
  
  // 构建 SDK 配置
  const sdkConfig = {
    // 使用流式输入以支持 MCP Tools
    prompt: generateMessages() as AsyncIterable<any>,
    options: {
      // 启用 resume 功能（SDK 会自动从 ~/.claude/projects/ 加载 session）
      ...(resume && { resume }),
      // 使用优先级配置的模型名称
      model: finalModelName,
      // 如果有自定义 API URL，则配置
      ...(finalApiUrl && { apiUrl: finalApiUrl }),
      allowedTools,
      settingSources: ["project" as const],
      // MCP Servers - 包含 Office 代码执行 Tool
      mcpServers: {
        "office": officeMcpServer
      },
      // 注意：不使用 permissionMode: 'plan'，因为它会导致所有工具都需要批准
      // 我们通过系统提示语实现"先规划后执行"的行为，更适合 Office 操作场景
      canUseTool: async (toolName: string, input: any, context: any) => {
        // MCP Tools 和 Office 工具都允许
        if (toolName.startsWith("mcp__") || toolName.startsWith("office_")) {
          return { behavior: "allow" as const };
        }
        return { behavior: "allow" as const };
      }
    }
  };
  
  // 如果有自定义 API Key，设置环境变量（临时覆盖）
  if (apiKey && apiKey.trim()) {
    process.env.ANTHROPIC_API_KEY = apiKey.trim();
  }

  // 日志输出 - 增强诊断信息
  console.log('[API/chat] Configuration:');
  console.log('  - API Key source:', apiKey?.trim() ? 'Frontend' : process.env.ANTHROPIC_API_KEY ? 'Environment' : 'None');
  console.log('  - API URL:', finalApiUrl || '(using default)');
  console.log('  - Model:', finalModelName);
  console.log('  - Mode:', mode);
  console.log('  - Host Type:', hostType);
  
  if (process.env.ANTHROPIC_MODEL) {
    console.log('[API/chat] Using custom model:', process.env.ANTHROPIC_MODEL);
  }
  if (process.env.ANTHROPIC_BASE_URL) {
    console.log('[API/chat] Using custom API base URL:', process.env.ANTHROPIC_BASE_URL);
  }
  if (resume) {
    console.log('[API/chat] Resuming session:', resume);
  } else {
    console.log('[API/chat] Starting new session (no resume)');
  }
  if (testMode) {
    console.log("[API/chat] Test mode request:", {
      testSessionId,
      testCaseId,
      hostType,
      promptPreview: prompt?.slice(0, 200),
      mode,
    });
  }
  
  let cancelled = false;

  const stream = new ReadableStream({
    async start(controller) {
      const safeEnqueue = (payload: string): boolean => {
        if (cancelled) return false;
        try {
          controller.enqueue(encoder.encode(payload));
          return true;
        } catch (e) {
          // 客户端断开或 controller 已关闭
          cancelled = true;
          return false;
        }
      };

      try {
        for await (const message of query(sdkConfig)) {
          if (cancelled) break;
          
          // 只从系统初始化消息中获取 session ID（官方推荐方式）
          if (message.type === 'system' && message.subtype === 'init' && message.session_id) {
            safeEnqueue(`event: session\ndata: ${JSON.stringify({ sessionId: message.session_id })}\n\n`);
            console.log('[API/chat] Session ID:', message.session_id);
            continue; // ✅ 跳过后续的通用消息转发，避免重复
          }

          // 检测 TodoWrite 工具调用
          if (message.type === 'assistant' && message.message?.content) {
            for (const content of message.message.content) {
              // 检测工具调用
              if (content.type === 'tool_use' && content.name === 'TodoWrite') {
                console.log('[API/chat] Detected TodoWrite tool call:', content.input);
                
                // 发送 todos 事件
                const input = content.input as any;
                if (input && input.todos && Array.isArray(input.todos)) {
                  // 确保每个任务都有唯一的 id
                  const todosWithIds = input.todos.map((todo: any, index: number) => ({
                    ...todo,
                    id: todo.id || `task-${Date.now()}-${index}`,
                    content: todo.content || `任务 ${index + 1}`,
                    status: todo.status || 'pending'
                  }));
                  
                  safeEnqueue(`event: todos\ndata: ${JSON.stringify({
                    todos: todosWithIds,
                    title: input.title || '任务规划',
                    objective: input.objective
                  })}\n\n`);
                  console.log('[API/chat] Sent todos event with', todosWithIds.length, 'tasks');
                }
                
                // 继续发送原始消息（用于显示"正在执行"）
                break;
              }
            }
          }
          
          // 检测工具执行结果中的 todos 数据（使用类型断言处理未定义的属性）
          if (message.type === 'result') {
            const msgAny = message as any;
            if (msgAny.tool_name === 'TodoWrite' && msgAny.result) {
              try {
                const result = typeof msgAny.result === 'string' 
                  ? JSON.parse(msgAny.result) 
                  : msgAny.result;
                
                if (result && result.todos && Array.isArray(result.todos)) {
                  // 确保每个任务都有唯一的 id
                  const todosWithIds = result.todos.map((todo: any, index: number) => ({
                    ...todo,
                    id: todo.id || `task-${Date.now()}-${index}`,
                    content: todo.content || `任务 ${index + 1}`,
                    status: todo.status || 'pending'
                  }));
                  
                  console.log('[API/chat] Detected todos in TodoWrite result:', todosWithIds.length, 'tasks');
                  safeEnqueue(`event: todos\ndata: ${JSON.stringify({
                    todos: todosWithIds,
                    title: result.title || '任务规划',
                    objective: result.objective
                  })}\n\n`);
                }
              } catch (e) {
                console.warn('[API/chat] Failed to parse TodoWrite result:', e);
              }
            }
          }
          
          // 检测任务更新（如果 SDK 支持，使用类型断言）
          const msgAny = message as any;
          if (msgAny.type === 'task_update' && msgAny.task_id) {
            safeEnqueue(`event: task_update\ndata: ${JSON.stringify({
              taskId: msgAny.task_id,
              status: msgAny.status,
              result: msgAny.result,
              error: msgAny.error
            })}\n\n`);
            console.log('[API/chat] Sent task_update event for task:', msgAny.task_id);
            // 不要 continue，让原始消息也被发送
          }

          // 转发 SDK 消息
          if (!safeEnqueue(`event: message\ndata: ${JSON.stringify(message)}\n\n`)) {
            break;
          }
        }
        
        // 发送完成事件
        safeEnqueue(`event: complete\ndata: {}\n\n`);
        
      } catch (error) {
        console.error("[API/chat] Error:", error);
        const errorMessage = error instanceof Error ? error.message : String(error);

        try {
          if (cancelled) return;
          
          // 检测各种会话相关错误和 SDK 崩溃
          const isSessionNotFound = 
            errorMessage.includes("No conversation found with session ID") ||
            errorMessage.includes("Claude Code process exited") ||
            errorMessage.includes("process exited with code 1");

          if (isSessionNotFound) {
            console.warn('[API/chat] Session error detected, consider creating new session');
            // 通知前端清除无效的 session ID
            safeEnqueue(
              `event: session_invalid\ndata: ${JSON.stringify({
                message: "会话错误，建议使用新会话..."
              })}\n\n`
            );
          }

          safeEnqueue(
            `event: message\ndata: ${JSON.stringify({
              type: "result",
              subtype: "error_during_execution",
              errors: [errorMessage]
            })}\n\n`
          );
        } catch (controllerError) {
          // Controller 已关闭，忽略错误
          console.error("[API/chat] Controller already closed:", controllerError);
        }
      } finally {
        try {
          if (!cancelled) controller.close();
        } catch (closeError) {
          // Controller 已关闭，忽略错误
        }
      }
    },
    cancel() {
      cancelled = true;
    },
  });

  return new Response(stream, {
    headers: {
      "Content-Type": "text/event-stream",
      "Cache-Control": "no-cache",
      "Connection": "keep-alive",
    },
  });
}
