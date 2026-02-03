export type MessageRole =
  | "user"
  | "assistant"
  | "plan"
  | "thought"
  | "status"
  | "action"
  | "observation"
  | "error"
  | "info"
  | "chart"
  | "task_list";

export interface ChatMessage {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: string;
  metadata?: {
    toolName?: string;
    toolInput?: unknown;
    toolResult?: unknown;
    plotUrls?: string[];
    fileId?: string;
    filename?: string;
    tasks?: TodoItem[];
    taskTitle?: string;
    taskObjective?: string;
    retryCount?: number;
    maxRetries?: number;
    errorType?: string;
    finalError?: boolean;
  };
  streaming?: boolean;
}

export interface Conversation {
  id: string;
  title: string;
  sessionId?: string;
  messages: ChatMessage[];
  streamingMessageId?: string;
}

export type TaskStatus = "pending" | "in_progress" | "completed" | "failed";

export interface TodoItem {
  id: string;
  content: string;
  status: TaskStatus;
  activeForm?: string;
  result?: string;
  error?: string;
}

export interface ChatRequestPayload {
  prompt: string;
  hostType: string;
  resume?: string;
  attachments?: string[];
  metadata?: Record<string, unknown>;
}

export interface SSEEvent {
  type: string;
  data: Record<string, unknown>;
}
