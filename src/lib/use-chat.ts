"use client";

import { useState, useCallback, useRef, useEffect } from "react";
import { ChatMessage, TodoItem, SSEEvent } from "@/types/chat";
import { OfficeHostType } from "./office/host-detector";
import { handleOfficeAction } from "./office/bridge-factory";
import { mapSDKMessageToUI } from "./message-mapper";
import { extractHiddenOfficeCode, removeHiddenOfficeCode, executeOfficeCode } from "./office/code-executor";
import { loadApiSettings } from "@/components/SettingsDialog";
import { buildErrorFeedback, buildFinalErrorMessage } from "./office/error-feedback-builder";

const DEBUG_CHAT =
  process.env.NODE_ENV === "development" ||
  (typeof process !== "undefined" && process.env?.NEXT_PUBLIC_DEBUG_CHAT === "1");

// 最大重试次数（用于旧的隐藏代码机制）
const MAX_RETRIES = 3;

// MCP Tool 执行配置
const MCP_POLL_INTERVAL = 200; // 轮询间隔（降级方案使用）
const MCP_SSE_RECONNECT_DELAY = 3000; // SSE 重连延迟（毫秒）
const MCP_SSE_ENABLED = true; // 是否启用 SSE（可配置）

interface EventHandlers {
  setMessages: React.Dispatch<React.SetStateAction<ChatMessage[]>>;
  setTodos: React.Dispatch<React.SetStateAction<TodoItem[]>>;
  setShowTodos: React.Dispatch<React.SetStateAction<boolean>>;
  onSessionIdChange: (sessionId: string) => void;
  hostType: OfficeHostType;
  retryStateRef: React.MutableRefObject<Map<string, number>>;
  latestSessionIdRef: React.MutableRefObject<string | null>;
}

interface UseChatOptions {
  sessionId?: string | null; // 从外部传入的 session ID
  onSessionIdChange?: (sessionId: string) => void; // 当 session ID 变化时的回调
}

// 上传文件信息接口
export interface UploadedFile {
  name: string;
  size: number;
  file: File;
  serverPath?: string;  // 服务器返回的相对路径
  fileId?: string;      // 文件唯一标识
}

export function useChat(hostType: OfficeHostType, options: UseChatOptions = {}) {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [todos, setTodos] = useState<TodoItem[]>([]);
  const [showTodos, setShowTodos] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  
  // 重试状态管理：跟踪每个用户消息的重试次数
  const retryStateRef = useRef<Map<string, number>>(new Map());
  
  // 使用 ref 跟踪最新的 sessionId（用于重试）
  const latestSessionIdRef = useRef<string | null>(options.sessionId || null);
  
  const abortControllerRef = useRef<AbortController | null>(null);
  
  // 从外部获取 sessionId
  const sessionId = options.sessionId;
  const onSessionIdChange = options.onSessionIdChange;
  
  // 同步更新 latestSessionIdRef
  useEffect(() => {
    latestSessionIdRef.current = sessionId || null;
  }, [sessionId]);
  
  // 已处理的 MCP 执行任务 ID（避免重复执行）
  const processedMcpTasksRef = useRef<Set<string>>(new Set());
  
  // SSE 连接状态（用于降级）
  const [useSSE, setUseSSE] = useState(MCP_SSE_ENABLED);
  
  // MCP Tool 执行：优先使用 SSE，失败时降级到轮询
  useEffect(() => {
    if (!isLoading) {
      // 清理已处理的任务 ID（会话结束后）
      processedMcpTasksRef.current.clear();
      return;
    }
    
    // 任务执行逻辑（SSE 和轮询共用）
    const executeTask = async (task: any) => {
      // 跳过已处理的任务
      if (processedMcpTasksRef.current.has(task.correlationId)) {
        return;
      }
      
      // 标记为已处理
      processedMcpTasksRef.current.add(task.correlationId);
      
      if (DEBUG_CHAT) {
        console.log('[useChat] MCP Tool task received:', task.correlationId, task.host);
      }
      
      // 显示执行中状态
      setMessages(prev => [...prev, {
        id: crypto.randomUUID(),
        role: "info",
        content: `🔄 正在执行 ${task.host.toUpperCase()} 代码...`,
        timestamp: new Date().toLocaleTimeString(),
      }]);
      
      // 执行代码
      try {
        const execStartTime = Date.now();
        const result = await executeOfficeCode(task.host, task.code);
        const execDuration = Date.now() - execStartTime;
        
        if (DEBUG_CHAT) {
          console.log('[useChat] MCP Tool execution result:', {
            correlationId: task.correlationId,
            success: result.success,
            duration: execDuration + 'ms',
            hasResult: result.result !== undefined,
            resultType: typeof result.result,
            errorType: result.error?.type,
          });
        }
        
        // 非调试模式下也记录关键信息
        if (!result.success) {
          console.error('[useChat] MCP Tool execution failed:', {
            host: task.host,
            errorType: result.error?.type,
            errorMessage: result.error?.message,
          });
        } else if (result.result === undefined && execDuration < 100) {
          console.warn('[useChat] MCP Tool execution succeeded but returned undefined in < 100ms - possible silent failure');
        }
        
        // 提交结果到服务端
        await fetch('/api/tool-result', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            correlationId: task.correlationId,
            source: 'mcp_tool',
            result: {
              success: result.success,
              data: result.result,
              error: result.error ? {
                type: result.error.type,
                code: result.error.code,
                message: result.error.message,
                stackTrace: result.error.stackTrace,
              } : undefined,
            }
          })
        });
        
        // 更新 UI 显示执行结果
        if (result.success) {
          // 移除"执行中"消息，因为 Agent 会返回成功信息
          setMessages(prev => prev.filter(msg => 
            !(msg.role === "info" && msg.content.includes("正在执行"))
          ));
        } else {
          // 执行失败，显示简短提示（Agent 会自动修复）
          setMessages(prev => {
            // 替换"执行中"消息为"修复中"
            return prev.map(msg => 
              msg.role === "info" && msg.content.includes("正在执行")
                ? { ...msg, content: `⚠️ 代码执行失败，AI 正在自动修复...` }
                : msg
            );
          });
        }
      } catch (execError) {
        console.error('[useChat] MCP Tool execution error:', execError);
        
        // 提交错误结果
        await fetch('/api/tool-result', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            correlationId: task.correlationId,
            source: 'mcp_tool',
            result: {
              success: false,
              error: {
                type: 'UnknownError',
                message: execError instanceof Error ? execError.message : String(execError),
              }
            }
          })
        });
      }
    };
    
    // 方案 1：SSE 推送
    if (useSSE) {
      let eventSource: EventSource | null = null;
      let reconnectTimer: NodeJS.Timeout | null = null;
      let reconnectAttempts = 0;
      const maxReconnectAttempts = 3;
      
      const connectSSE = () => {
        try {
          // 建立 SSE 连接
          eventSource = new EventSource('/api/task-stream');
          
          // 连接建立
          eventSource.addEventListener('open', () => {
            console.log('[useChat/SSE] Connected successfully');
            reconnectAttempts = 0; // 重置重连次数
          });
          
          // 接收消息
          eventSource.addEventListener('message', async (event) => {
            try {
              const data = JSON.parse(event.data);
              
              if (data.type === 'connected') {
                console.log('[useChat/SSE] Connection established');
                return;
              }
              
              if (data.type === 'task') {
                console.log('[useChat/SSE] Task received:', data.correlationId);
                await executeTask(data);
              }
            } catch (error) {
              console.error('[useChat/SSE] Message parse error:', error);
            }
          });
          
          // 连接错误
          eventSource.addEventListener('error', (error) => {
            console.error('[useChat/SSE] Connection error:', error);
            eventSource?.close();
            eventSource = null;
            
            reconnectAttempts++;
            
            // 超过最大重连次数，降级到轮询
            if (reconnectAttempts >= maxReconnectAttempts) {
              console.warn(`[useChat/SSE] Max reconnect attempts (${maxReconnectAttempts}) reached, falling back to polling`);
              setUseSSE(false);
              return;
            }
            
            // 重连
            reconnectTimer = setTimeout(() => {
              console.log(`[useChat/SSE] Reconnecting... (attempt ${reconnectAttempts + 1}/${maxReconnectAttempts})`);
              connectSSE();
            }, MCP_SSE_RECONNECT_DELAY);
          });
        } catch (error) {
          console.error('[useChat/SSE] Connection setup error:', error);
          // 降级到轮询
          setUseSSE(false);
        }
      };
      
      // 启动 SSE 连接
      connectSSE();
      
      // 清理函数
      return () => {
        if (reconnectTimer) clearTimeout(reconnectTimer);
        if (eventSource) {
          eventSource.close();
          console.log('[useChat/SSE] Connection closed');
        }
      };
    }
    
    // 方案 2：轮询（降级方案）
    let cancelled = false;
    
    const pollForMcpTasks = async () => {
      console.log('[useChat/Polling] Started (SSE unavailable)');
      
      while (!cancelled) {
        try {
          const response = await fetch('/api/tool-result?action=pending_executions');
          if (!response.ok) {
            await new Promise(resolve => setTimeout(resolve, MCP_POLL_INTERVAL));
            continue;
          }
          
          const data = await response.json();
          const executions = data.executions || [];
          
          for (const task of executions) {
            await executeTask(task);
          }
        } catch (pollError) {
          if (DEBUG_CHAT) {
            console.warn('[useChat/Polling] Error:', pollError);
          }
        }
        
        // 等待下一次轮询
        await new Promise(resolve => setTimeout(resolve, MCP_POLL_INTERVAL));
      }
    };
    
    // 启动轮询
    pollForMcpTasks();
    
    return () => {
      cancelled = true;
    };
  }, [isLoading, hostType, useSSE]);

  const sendMessage = useCallback(async (content: string, mode: 'Agent' | 'Plan' = 'Agent', uploadedFiles?: UploadedFile[]) => {
    // 增强 prompt，包含上传文件信息
    let finalPrompt = content;
    
    if (uploadedFiles && uploadedFiles.length > 0) {
      const fileInfo = uploadedFiles
        .filter(f => f.serverPath) // 只包含已成功上传的文件
        .map(f => `- ${f.name} (路径: ${f.serverPath})`)
        .join('\n');
      
      if (fileInfo) {
        finalPrompt = `${content}\n\n[已上传文件]\n${fileInfo}\n\n请使用 Read 工具读取这些文件。`;
      }
    }
    
    // 添加用户消息
    const userMessage: ChatMessage = {
      id: crypto.randomUUID(),
      role: "user",
      content,
      timestamp: new Date().toLocaleTimeString(),
    };
    
    // 添加思考中占位符
    const thinkingPlaceholderId = crypto.randomUUID();
    const thinkingMessage: ChatMessage = {
      id: thinkingPlaceholderId,
      role: "assistant",
      content: "",
      timestamp: "",
      streaming: true,
    };
    
    setMessages(prev => [...prev, userMessage, thinkingMessage]);
    setIsLoading(true);
    
    // 创建 AbortController
    abortControllerRef.current = new AbortController();

    try {
      // 从 localStorage 读取 API 配置
      const apiSettings = loadApiSettings();
      
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ 
          prompt: finalPrompt,  // 使用增强后的 prompt
          hostType,
          mode,
          ...(sessionId && { resume: sessionId }),
          // 传递 API 配置（如果有）
          ...(apiSettings?.apiKey && { apiKey: apiSettings.apiKey }),
          ...(apiSettings?.apiUrl && { apiUrl: apiSettings.apiUrl }),
          ...(apiSettings?.modelName && { modelName: apiSettings.modelName }),
        }),
        signal: abortControllerRef.current.signal,
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const reader = response.body!.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      let currentStreamingId: string | null = null;
      let hasReceivedFirstMessage = false;

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        buffer += decoder.decode(value, { stream: true });
        
        // 解析 SSE 事件
        const { parsed, remaining } = parseSSEBuffer(buffer);
        buffer = remaining;

        for (const event of parsed) {
          if (DEBUG_CHAT) {
            const summary =
              event.type === "message"
                ? `contentLen=${typeof (event.data as { content?: string }).content === "string" ? (event.data as { content: string }).content.length : 0}`
                : event.type === "office_action"
                  ? `action=${(event.data as { action?: string }).action} correlationId=${(event.data as { correlationId?: string }).correlationId} payloadKeys=${Object.keys((event.data as { payload?: Record<string, unknown> }).payload ?? {}).join(",")}`
                  : "";
            console.log("[useChat] SSE event", event.type, new Date().toISOString(), summary || event.data);
          }
          
          // 收到第一条消息时，移除思考占位符
          if (event.type === "message" && !hasReceivedFirstMessage) {
            hasReceivedFirstMessage = true;
            setMessages(prev => prev.filter(msg => msg.id !== thinkingPlaceholderId));
          }
          
          await handleSSEEvent(event, {
            setMessages,
            setTodos,
            setShowTodos,
            onSessionIdChange: onSessionIdChange || (() => {}),
            hostType,
            retryStateRef,
            latestSessionIdRef,
          }, currentStreamingId, (id) => { currentStreamingId = id; });
        }
      }
    } catch (error) {
      // 移除思考占位符
      setMessages(prev => prev.filter(msg => msg.id !== thinkingPlaceholderId));
      
      if ((error as Error).name !== "AbortError") {
        setMessages(prev => [...prev, {
          id: crypto.randomUUID(),
          role: "error",
          content: `请求失败: ${(error as Error).message}`,
          timestamp: new Date().toLocaleTimeString(),
        }]);
      }
    } finally {
      setIsLoading(false);
      // 清理 AbortController 引用
      abortControllerRef.current = null;
    }
  }, [sessionId, hostType, onSessionIdChange]);

  const abort = useCallback(() => {
    abortControllerRef.current?.abort();
  }, []);

  // 加载消息到当前会话
  const loadMessages = useCallback((newMessages: ChatMessage[]) => {
    setMessages(newMessages);
  }, []);

  // 清空当前会话的消息（但不影响 sessionId）
  const clearMessages = useCallback(() => {
    setMessages([]);
    setTodos([]);
    // 清理重试状态
    retryStateRef.current.clear();
  }, []);

  return {
    messages,
    todos,
    showTodos,
    isLoading,
    sendMessage,
    abort,
    setShowTodos,
    loadMessages,
    clearMessages,
  };
}

// 跟踪已执行的代码块哈希，防止重复执行
const executedCodeHashes = new Set<string>();

// 简单的字符串哈希函数
function hashCode(str: string): string {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = (hash << 5) - hash + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return hash.toString();
}

// SSE 事件处理
async function handleSSEEvent(
  event: SSEEvent, 
  handlers: EventHandlers,
  currentStreamingId: string | null,
  setCurrentStreamingId: (id: string | null) => void
) {
  const { setMessages, setTodos, setShowTodos, onSessionIdChange, hostType, retryStateRef, latestSessionIdRef } = handlers;

  switch (event.type) {
    case "session":
      // 通知外部 session ID 变化
      const newSessionId = event.data.sessionId as string;
      console.log("[useChat] Session ID from SDK:", newSessionId);
      // 更新最新的 session ID ref
      latestSessionIdRef.current = newSessionId;
      onSessionIdChange(newSessionId);
      break;

    case "session_invalid":
      // 清除无效的 session ID
      console.warn("[useChat] Session expired, clearing...");
      onSessionIdChange("");
      // 显示提示消息
      setMessages(prev => [...prev, {
        id: crypto.randomUUID(),
        role: "info",
        content: "会话已过期，已自动开始新对话。请重新发送您的消息。",
        timestamp: new Date().toLocaleTimeString(),
      }]);
      break;

    case "message": {
      const uiMessage = mapSDKMessageToUI(event.data as any);
      if (uiMessage) {
        // 如果是流式传输且为 assistant 消息，暂时不执行代码，等待完整消息
        // 这样可以避免在流式传输过程中执行不完整的代码，或者重复执行
        const shouldExecuteCode = !uiMessage.streaming && uiMessage.role === 'assistant';

        // 检测并自动执行隐藏的 Office 代码
        if (shouldExecuteCode && uiMessage.content && typeof uiMessage.content === 'string') {
          const codeBlocks = extractHiddenOfficeCode(uiMessage.content);
          
          // 从显示内容中移除隐藏代码
          if (codeBlocks.length > 0) {
            uiMessage.content = removeHiddenOfficeCode(uiMessage.content);
            
            // 异步执行代码（不阻塞消息显示）
            for (const block of codeBlocks) {
              const codeHash = hashCode(block.code);
              
              // 检查是否已执行过此代码
              if (executedCodeHashes.has(codeHash)) {
                console.log('[useChat] Skipping duplicate code execution:', codeHash);
                continue;
              }
              
              executedCodeHashes.add(codeHash);
              
              executeOfficeCode(block.host, block.code).then(async result => {
                if (!result.success) {
                  // 执行失败，检查是否可以重试
                  console.error('[useChat] Office code execution failed:', result.error);
                  
                  // 获取当前消息的原始用户输入（需要从消息历史中推断）
                  // 使用代码内容作为重试键
                  const retryKey = codeHash;
                  const currentRetries = retryStateRef.current.get(retryKey) || 0;
                  
                  if (currentRetries < MAX_RETRIES && result.error) {
                    // 更新重试计数
                    retryStateRef.current.set(retryKey, currentRetries + 1);
                    
                    // 构建错误反馈
                    const errorFeedback = buildErrorFeedback(
                      result.error,
                      block.code,
                      currentRetries + 1,
                      MAX_RETRIES
                    );
                    
                    console.log(`[useChat] Auto-retrying (${currentRetries + 1}/${MAX_RETRIES})...`);
                    
                    // 显示重试提示消息
                    setMessages(prev => [...prev, {
                      id: crypto.randomUUID(),
                      role: "error",
                      content: `⚠️ 代码执行失败，正在自动重试（${currentRetries + 1}/${MAX_RETRIES}）...`,
                      timestamp: new Date().toLocaleTimeString(),
                      metadata: {
                        retryCount: currentRetries + 1,
                        maxRetries: MAX_RETRIES,
                        errorType: result.error?.type,
                      }
                    }]);
                    
                    // 自动发送错误反馈给 Agent（触发重新生成代码）
                    // 注意：这里直接调用 sendMessage 会导致递归，需要通过 API 直接发送
                    try {
                      const apiSettings = loadApiSettings();
                      const currentSessionId = latestSessionIdRef.current;
                      
                      console.log(`[useChat] Retry request with sessionId:`, currentSessionId);
                      
                      const response = await fetch("/api/chat", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ 
                          prompt: errorFeedback,
                          hostType,
                          mode: 'Agent',
                          ...(currentSessionId && { resume: currentSessionId }),
                          ...(apiSettings?.apiKey && { apiKey: apiSettings.apiKey }),
                          ...(apiSettings?.apiUrl && { apiUrl: apiSettings.apiUrl }),
                          ...(apiSettings?.modelName && { modelName: apiSettings.modelName }),
                        }),
                      });

                      if (response.ok) {
                        // 处理重试响应（复用现有的 SSE 处理逻辑）
                        const reader = response.body!.getReader();
                        const decoder = new TextDecoder();
                        let buffer = "";
                        let currentStreamingId: string | null = null;

                        while (true) {
                          const { done, value } = await reader.read();
                          if (done) break;

                          buffer += decoder.decode(value, { stream: true });
                          const { parsed, remaining } = parseSSEBuffer(buffer);
                          buffer = remaining;

                          for (const event of parsed) {
                            await handleSSEEvent(event, {
                              setMessages,
                              setTodos,
                              setShowTodos,
                              onSessionIdChange,
                              hostType,
                              retryStateRef,
                              latestSessionIdRef,
                            }, currentStreamingId, (id) => { currentStreamingId = id; });
                          }
                        }
                      }
                    } catch (retryError) {
                      console.error('[useChat] Retry request failed:', retryError);
                      setMessages(prev => [...prev, {
                        id: crypto.randomUUID(),
                        role: "error",
                        content: `重试请求失败: ${retryError instanceof Error ? retryError.message : String(retryError)}`,
                        timestamp: new Date().toLocaleTimeString(),
                      }]);
                    }
                  } else {
                    // 重试次数用尽或无错误信息，显示最终错误
                    const finalMessage = result.error 
                      ? buildFinalErrorMessage(result.error, block.code, currentRetries)
                      : `❌ 操作失败：${result.error}\n\n提示：请确认 ${block.host.toUpperCase()} 文档已打开。`;
                    
                    setMessages(prev => [...prev, {
                      id: crypto.randomUUID(),
                      role: "error",
                      content: finalMessage,
                      timestamp: new Date().toLocaleTimeString(),
                      metadata: {
                        retryCount: currentRetries,
                        maxRetries: MAX_RETRIES,
                        finalError: true,
                      }
                    }]);
                    
                    // 清理重试状态
                    retryStateRef.current.delete(retryKey);
                  }
                } else {
                  // 成功执行，清理重试状态
                  retryStateRef.current.delete(codeHash);
                  // 成功时清理 hash，允许未来相同的操作再次执行（可选）
                  // setTimeout(() => executedCodeHashes.delete(codeHash), 5000);
                }
              });
            }
          }
        }
        
        const normalize = (s: string) => s.replace(/\r\n/g, "\n").trim();
        const isSameOrContained = (a: string, b: string) => {
          const na = normalize(a);
          const nb = normalize(b);
          if (!na || !nb) return false;
          if (na === nb) return true;
          return na.includes(nb) || nb.includes(na);
        };

        if (uiMessage.streaming) {
          // 流式消息：尽量只更新同一条消息；如果角色发生变化（如 thought -> assistant），创建新消息避免混杂
          if (currentStreamingId) {
            setMessages(prev => {
              const existing = prev.find(m => m.id === currentStreamingId);
              if (existing && existing.role === uiMessage.role) {
                return prev.map(msg =>
                  msg.id === currentStreamingId
                    ? { ...msg, content: msg.content + (uiMessage.content || "") }
                    : msg
                );
              }
              const newId = crypto.randomUUID();
              setCurrentStreamingId(newId);
              return [
                ...prev,
                {
                  ...uiMessage,
                  id: newId,
                  timestamp: new Date().toLocaleTimeString(),
                } as ChatMessage,
              ];
            });
          } else {
            const newId = crypto.randomUUID();
            setCurrentStreamingId(newId);
            setMessages(prev => [
              ...prev,
              {
                ...uiMessage,
                id: newId,
                timestamp: new Date().toLocaleTimeString(),
              } as ChatMessage,
            ]);
          }
        } else {
          // 非流式消息（通常是完整消息或结束消息）
          if (currentStreamingId) {
            if (uiMessage.role === "assistant") {
              // 将最终 assistant/result 覆盖到同一条流式消息上；不要在这里清空 currentStreamingId（等 complete）
              setMessages(prev =>
                prev.map(msg =>
                  msg.id === currentStreamingId ? ({ ...msg, ...uiMessage, streaming: false } as ChatMessage) : msg
                )
              );
            } else {
              // tool/action 等消息：独立追加，但保留 currentStreamingId，避免打断后续 final 更新
              setMessages(prev => [
                ...prev,
                {
                  ...uiMessage,
                  id: crypto.randomUUID(),
                  timestamp: new Date().toLocaleTimeString(),
                } as ChatMessage,
              ]);
            }
          } else {
            // 没有流式消息：追加，但做内容级去重/合并，避免 assistant + result 双份显示
            setMessages(prev => {
              const nextContent = typeof uiMessage.content === "string" ? uiMessage.content : "";
              const last = prev[prev.length - 1];
              if (last && last.role === uiMessage.role && isSameOrContained(last.content, nextContent)) {
                // 如果新内容更长，用更长的覆盖（避免“短/长”两条都显示）
                const lastNorm = normalize(last.content);
                const nextNorm = normalize(nextContent);
                if (nextNorm.length > lastNorm.length) {
                  return prev.map((m, idx) =>
                    idx === prev.length - 1 ? ({ ...m, ...uiMessage, content: nextContent } as ChatMessage) : m
                  );
                }
                return prev;
              }
              return [
                ...prev,
                {
                  ...uiMessage,
                  id: crypto.randomUUID(),
                  timestamp: new Date().toLocaleTimeString(),
                } as ChatMessage,
              ];
            });
          }
        }
      }
      break;
    }

    case "todos": {
      const todos = event.data.todos as TodoItem[];
      
      if (DEBUG_CHAT) {
        console.log('[useChat] Received todos event:', {
          count: todos?.length,
          todos: todos,
          title: event.data.title,
          objective: event.data.objective
        });
      }
      
      if (!todos || !Array.isArray(todos) || todos.length === 0) {
        console.warn('[useChat] Invalid todos data:', event.data);
        break;
      }
      
      setTodos(todos);
      setShowTodos(true);
      
      // 检查是否已存在任务列表消息，如果存在则更新，否则创建新的
      setMessages(prev => {
        const existingTaskListIndex = prev.findIndex(msg => msg.role === "task_list");
        
        if (existingTaskListIndex !== -1) {
          // 更新已存在的任务列表消息
          if (DEBUG_CHAT) {
            console.log('[useChat] Updating existing task list message at index:', existingTaskListIndex);
          }
          return prev.map((msg, index) =>
            index === existingTaskListIndex
              ? {
                  ...msg,
                  metadata: {
                    ...msg.metadata,
                    tasks: todos,
                    taskTitle: (event.data.title as string) || msg.metadata?.taskTitle || "任务规划",
                    taskObjective: (event.data.objective as string) || msg.metadata?.taskObjective,
                  },
                }
              : msg
          );
        } else {
          // 创建新的任务列表消息
          const taskListMessage: ChatMessage = {
            id: crypto.randomUUID(),
            role: "task_list",
            content: "",
            timestamp: new Date().toLocaleTimeString(),
            metadata: {
              tasks: todos,
              taskTitle: (event.data.title as string) || "任务规划",
              taskObjective: event.data.objective as string,
            },
          };
          
          if (DEBUG_CHAT) {
            console.log('[useChat] Creating new task list message:', taskListMessage);
          }
          
          return [...prev, taskListMessage];
        }
      });
      break;
    }

    case "task_update": {
      const taskId = event.data.taskId as string;
      const status = event.data.status as TodoItem["status"];
      const result = event.data.result as string;
      const error = event.data.error as string;
      
      // 更新 todos 状态（用于 TodoPanel）
      setTodos(prev => prev.map(todo => 
        todo.id === taskId
          ? { ...todo, status, result, error }
          : todo
      ));
      
      // 更新消息流中最新的任务列表消息（从后往前找）
      setMessages(prev => {
        const messages = [...prev];
        // 从后往前找到第一个任务列表消息
        for (let i = messages.length - 1; i >= 0; i--) {
          if (messages[i].role === "task_list" && messages[i].metadata?.tasks) {
            messages[i] = {
              ...messages[i],
              metadata: {
                ...messages[i].metadata,
                tasks: messages[i].metadata!.tasks!.map(task =>
                  task.id === taskId
                    ? { ...task, status, result, error }
                    : task
                ),
              },
            };
            break; // 只更新最新的任务列表消息
          }
        }
        return messages;
      });
      break;
    }

    case "office_action": {
      const action = event.data.action as string;
      const payload = event.data.payload as Record<string, unknown>;
      const correlationId = event.data.correlationId as string | undefined;
      const startMs = Date.now();
      const result = await handleOfficeAction(action, payload, hostType);
      const durationMs = Date.now() - startMs;
      const resultSummary =
        result && typeof result === "object" && "success" in result
          ? `success=${(result as { success?: boolean }).success} error=${(result as { error?: string }).error ?? "—"}`
          : String(result);
      if (DEBUG_CHAT) {
        console.log("[useChat] office_action done", { action, correlationId, resultSummary, durationMs });
      }
      if (DEBUG_CHAT) {
        console.log("[useChat] tool-result POST", correlationId, resultSummary);
      }
      await fetch("/api/tool-result", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ correlationId, result }),
      });
      break;
    }

    case "complete":
      setCurrentStreamingId(null);
      // 可选：清除 executedCodeHashes，但这可能会导致如果快速发送相同请求时重复执行
      // executedCodeHashes.clear();
      break;
  }
}

// SSE 缓冲区解析
function parseSSEBuffer(buffer: string): { parsed: SSEEvent[]; remaining: string } {
  const events: SSEEvent[] = [];
  const lines = buffer.split("\n\n");
  
  // 最后一个可能是不完整的事件
  const remaining = lines.pop() || "";
  
  for (const chunk of lines) {
    if (!chunk.trim()) continue;
    
    let eventType = "message";
    let data = "";
    
    for (const line of chunk.split("\n")) {
      if (line.startsWith("event:")) {
        eventType = line.slice(6).trim();
      } else if (line.startsWith("data:")) {
        data = line.slice(5).trim();
      }
    }
    
    if (data) {
      try {
        events.push({ type: eventType, data: JSON.parse(data) });
      } catch (e) {
        console.error("[parseSSE] Failed to parse:", data, e);
      }
    }
  }
  
  return { parsed: events, remaining };
}
