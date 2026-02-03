import { ChatMessage, MessageRole } from "@/types/chat";

interface SDKMessage {
  type: string;
  subtype?: string;
  session_id?: string;
  uuid?: string;
  message?: {
    content?: Array<{
      type: string;
      text?: string;
      name?: string;
      input?: unknown;
    }>;
  };
  thinking?: string;
  tool_name?: string;
  input?: unknown;
  result?: string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  [key: string]: any;
}

/**
 * 将 Claude Agent SDK 消息映射到前端 UI 消息格式
 */
export function mapSDKMessageToUI(sdkMessage: SDKMessage): Partial<ChatMessage> | null {
  switch (sdkMessage.type) {
    // 系统消息
    case "system":
      if (sdkMessage.subtype === "init") {
        return { 
          role: "status" as MessageRole, 
          content: "会话已初始化",
        };
      }
      return null;

    // 助手消息
    case "assistant": {
      const content = sdkMessage.message?.content;
      if (!content || content.length === 0) return null;
      
      // 查找文本内容
      const textContent = content.find(c => c.type === "text");
      if (textContent?.text) {
        return {
          role: "assistant" as MessageRole,
          content: textContent.text,
        };
      }
      
      // 查找工具调用
      const toolUse = content.find(c => c.type === "tool_use");
      if (toolUse) {
        // 过滤掉 TodoWrite 工具调用显示（它会以任务列表形式显示）
        if (toolUse.name === "TodoWrite") {
          return null;
        }
        
        return {
          role: "action" as MessageRole,
          content: `正在执行: ${toolUse.name}`,
          metadata: {
            toolName: toolUse.name,
            toolInput: toolUse.input,
          },
        };
      }
      
      return null;
    }

    // 流式事件
    case "stream_event": {
      // 处理流式文本
      if (sdkMessage.event?.type === "content_block_delta") {
        const delta = sdkMessage.event?.delta;
        if (delta?.type === "text_delta" && delta?.text) {
          return {
            role: "assistant" as MessageRole,
            content: delta.text,
            streaming: true,
          };
        }
        // 思考过程
        if (delta?.type === "thinking_delta" && delta?.thinking) {
          return {
            role: "thought" as MessageRole,
            content: delta.thinking,
            streaming: true,
          };
        }
      }
      return null;
    }

    // 结果消息
    case "result": {
      if (sdkMessage.subtype === "success" && sdkMessage.result) {
        return {
          role: "assistant" as MessageRole,
          content: sdkMessage.result,
        };
      }
      if (sdkMessage.subtype?.startsWith("error")) {
        return {
          role: "error" as MessageRole,
          content: sdkMessage.errors?.join("\n") || "执行出错",
        };
      }
      return null;
    }

    default:
      return null;
  }
}

/**
 * 格式化工具执行结果
 */
export function formatToolResult(result: unknown): string {
  if (typeof result === "string") {
    return result;
  }
  if (result && typeof result === "object") {
    // @ts-expect-error check for error property
    if (result.error) {
      // @ts-expect-error access error property
      return `错误: ${result.error}`;
    }
    // @ts-expect-error check for success property
    if (result.success !== undefined) {
      // @ts-expect-error access success property
      return result.success ? "执行成功" : "执行失败";
    }
    return JSON.stringify(result, null, 2);
  }
  return String(result);
}
