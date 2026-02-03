"use client";

import { useState, useRef, useEffect, useCallback, useMemo } from "react";
import { Plus, Clock, Settings } from "lucide-react";
import { ChatInput } from "./ChatInput";
import { MessageBubble } from "./MessageBubble";
import { ChatTab } from "./ChatTab";
import { HistoryPanel } from "./HistoryPanel";
import { OfficeTestPanel } from "./OfficeTestPanel";
import { SettingsDialog } from "./SettingsDialog";
import { useChat } from "@/lib/use-chat";
import { OfficeHostType } from "@/lib/office/host-detector";
import { OfficeBridge } from "@/lib/office/bridge-factory";
import { ChatMessage } from "@/types/chat";
import { loadSessions, saveSessions, deleteSession, migrateLegacySession } from "@/lib/session-storage";
import type { HostType as TestHostType } from "../../tests/office-skills/test-runner/types";

// 深度比较消息数组是否相等
function areMessagesEqual(a: ChatMessage[], b: ChatMessage[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (
      a[i].id !== b[i].id ||
      a[i].content !== b[i].content ||
      a[i].role !== b[i].role ||
      a[i].streaming !== b[i].streaming
    ) {
      return false;
    }
  }
  return true;
}

// 创建新会话
function createNewSession(): ChatSession {
  return {
    id: Date.now().toString(),
    name: `Chat ${Date.now()}`,
    messages: [],
    timestamp: new Date().toISOString(),
  };
}

interface SidebarProps {
  hostType: OfficeHostType;
  bridge: OfficeBridge | null;
}

interface ChatSession {
  id: string;
  name: string;
  sessionId?: string; // 后端的 session ID
  messages: ChatMessage[];
  timestamp: string; // ISO 格式的时间戳
}

interface HistoryItem {
  id: string;
  name: string;
  timestamp: string;
  preview: string;
}

export function Sidebar({ hostType, bridge }: SidebarProps) {
  const enableTestPanel = process.env.NEXT_PUBLIC_ENABLE_TEST_PANEL === "1";
  const [activeView, setActiveView] = useState<"chat" | "tests">("chat");
  const normalizedTestHostType: TestHostType =
    hostType === "excel" || hostType === "word" || hostType === "powerpoint" ? hostType : "excel";

  // 初始化：从 localStorage 加载会话或创建默认会话
  const [sessions, setSessions] = useState<ChatSession[]>(() => {
    const stored = loadSessions();
    if (stored.length > 0) {
      return stored;
    }
    // 如果没有存储的会话，创建一个默认会话
    const defaultSession = createNewSession();
    // 尝试迁移旧的 sessionId
    const legacySessionId = migrateLegacySession();
    if (legacySessionId) {
      defaultSession.sessionId = legacySessionId;
    }
    return [defaultSession];
  });
  
  const [activeSessionId, setActiveSessionId] = useState(() => sessions[0].id);
  const [isHistoryOpen, setIsHistoryOpen] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);

  const activeSession = sessions.find(s => s.id === activeSessionId);
  const currentSessionId = activeSession?.sessionId;

  // 当后端返回新的 sessionId 时更新当前会话
  const handleSessionIdChange = useCallback((newSessionId: string) => {
    setSessions(prev => {
      const updated = prev.map(session => {
        if (session.id === activeSessionId) {
          return {
            ...session,
            sessionId: newSessionId || undefined,
          };
        }
        return session;
      });
      // 手动保存到 localStorage
      saveSessions(updated);
      return updated;
    });
  }, [activeSessionId]);

  const {
    messages: chatMessages,
    isLoading,
    sendMessage: sendChatMessage,
    abort: abortChatMessage,
    loadMessages,
    clearMessages,
  } = useChat(hostType, {
    sessionId: currentSessionId,
    onSessionIdChange: handleSessionIdChange,
  });

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const lastLoadedRef = useRef<{ sessionId: string; length: number; lastId: string | null }>({ sessionId: "", length: -1, lastId: null });
  const lastSyncedMessagesRef = useRef<ChatMessage[]>([]);
  const sessionsRef = useRef(sessions);
  const isInitializedRef = useRef(false);

  const isDev = process.env.NODE_ENV === "development";

  // 保持 sessionsRef 同步
  useEffect(() => {
    sessionsRef.current = sessions;
  }, [sessions]);

  // 初始化：加载第一个会话的消息
  useEffect(() => {
    if (!isInitializedRef.current && sessions.length > 0) {
      isInitializedRef.current = true;
      const firstSession = sessions[0];
      if (firstSession.messages.length > 0) {
        lastLoadedRef.current = { 
          sessionId: firstSession.id, 
          length: firstSession.messages.length, 
          lastId: firstSession.messages[firstSession.messages.length - 1].id 
        };
        lastSyncedMessagesRef.current = firstSession.messages;
        loadMessages(firstSession.messages);
        if (isDev) {
          console.log("[Sidebar] initialized with messages from session", firstSession.id);
        }
      }
    }
  }, [sessions, loadMessages, isDev]);

  // 当切换会话时，加载该会话的消息到 useChat（只依赖 activeSessionId，避免循环）
  useEffect(() => {
    // 跳过初始化阶段
    if (!isInitializedRef.current) return;
    
    const session = sessionsRef.current.find(s => s.id === activeSessionId);
    if (!session) return;
    const msgs = session.messages;
    const len = msgs.length;
    const lastId = len > 0 ? msgs[len - 1].id : null;
    const prev = lastLoadedRef.current;
    
    // 检查是否需要加载
    if (prev.sessionId === activeSessionId && prev.length === len && prev.lastId === lastId) {
      return;
    }
    
    lastLoadedRef.current = { sessionId: activeSessionId, length: len, lastId };
    lastSyncedMessagesRef.current = msgs;
    
    if (isDev) {
      console.log("[Sidebar] effect load ran, activeSessionId=", activeSessionId);
    }
    loadMessages(msgs);
  }, [activeSessionId, loadMessages, isDev]);

  // 同步 useChat 的消息到当前活跃会话（使用深度比较避免循环）
  useEffect(() => {
    // 跳过初始化阶段的同步
    if (!isInitializedRef.current) return;
    
    // 深度比较：只有消息真正发生变化时才更新
    if (areMessagesEqual(chatMessages, lastSyncedMessagesRef.current)) {
      return;
    }
    
    lastSyncedMessagesRef.current = chatMessages;
    
    if (isDev) {
      console.log("[Sidebar] effect sync ran, activeSessionId=", activeSessionId, "messagesCount=", chatMessages.length);
    }
    
    setSessions(prev => {
      const updated = prev.map(session => {
        if (session.id === activeSessionId) {
          // 再次检查是否真的不同，避免不必要的更新
          if (areMessagesEqual(session.messages, chatMessages)) {
            return session;
          }
          
          const updatedSession = {
            ...session,
            messages: chatMessages,
            timestamp: new Date().toISOString(), // 更新时间戳
          };
          
          // 检查是否需要生成标题
          if (shouldGenerateTitle(updatedSession)) {
            const userMessages = updatedSession.messages.filter(m => m.role === 'user');
            if (userMessages.length > 0) {
              const newTitle = generateChatTitle(userMessages[0].content);
              if (isDev) {
                console.log("[Sidebar] Generated title:", newTitle);
              }
              updatedSession.name = newTitle;
            }
          }
          
          return updatedSession;
        }
        return session;
      });
      // 手动保存到 localStorage
      saveSessions(updated);
      return updated;
    });
  }, [chatMessages, activeSessionId, isDev]);

  // 自动滚动到底部
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [activeSession?.messages]);

  const handleSendMessage = async (content: string, mode: 'Agent' | 'Plan', uploadedFiles?: any[]) => {
    // 调用真正的 API，传递文件信息
    await sendChatMessage(content, mode, uploadedFiles);
  };

  const handleCreateNewChat = () => {
    const newSession = createNewSession();
    newSession.name = `Chat ${sessions.length + 1}`;
    const updatedSessions = [...sessions, newSession];
    setSessions(updatedSessions);
    // 手动保存到 localStorage
    saveSessions(updatedSessions);
    // 清空消息（新会话没有 sessionId）
    clearMessages();
    setActiveSessionId(newSession.id);
  };

  const handleCloseChat = (sessionId: string) => {
    if (sessions.length === 1) return; // Keep at least one session
    
    setSessions(prev => prev.filter(s => s.id !== sessionId));
    
    if (activeSessionId === sessionId) {
      const remainingSessions = sessions.filter(s => s.id !== sessionId);
      setActiveSessionId(remainingSessions[0].id);
    }
  };

  const handleSelectHistory = (historyId: string) => {
    // 检查是否已经打开
    const existingSession = sessions.find(s => s.id === historyId);
    if (existingSession) {
      setActiveSessionId(historyId);
      return;
    }

    // 从 localStorage 加载历史会话
    const allSessions = loadSessions();
    const historySession = allSessions.find(s => s.id === historyId);
    
    if (historySession) {
      const updatedSessions = [...sessions, historySession];
      setSessions(updatedSessions);
      // 手动保存到 localStorage（虽然数据已存在，但保持一致性）
      saveSessions(updatedSessions);
      setActiveSessionId(historyId);
    }
  };

  const handleDeleteHistory = (historyId: string) => {
    // 从 localStorage 删除
    deleteSession(historyId);
    
    // 如果该会话当前已打开，也从 sessions 中移除
    setSessions(prev => {
      const filtered = prev.filter(s => s.id !== historyId);
      
      // 如果删除的是当前活动会话，切换到第一个会话
      if (historyId === activeSessionId && filtered.length > 0) {
        setActiveSessionId(filtered[0].id);
      }
      
      // 如果删除后没有会话了，创建一个新的
      if (filtered.length === 0) {
        const newSession = createNewSession();
        setActiveSessionId(newSession.id);
        // 保存新创建的会话到 localStorage
        saveSessions([newSession]);
        return [newSession];
      }
      
      return filtered;
    });
  };

  // 动态生成历史对话列表
  const historyItems = useMemo(() => {
    // 仅在历史面板打开时计算，优化性能
    if (!isHistoryOpen) return [];

    // 加载所有存储的会话
    const storageSessions = loadSessions();
    
    // 获取当前打开的会话 ID 集合
    const openSessionIds = new Set(sessions.map(s => s.id));
    
    // 从存储中获取未打开的会话
    const closedSessions = storageSessions.filter(s => !openSessionIds.has(s.id));
    
    // 合并当前打开的会话和未打开的历史会话
    // 注意：使用 sessions 中的最新状态作为打开会话的数据源
    const allSessions = [...sessions, ...closedSessions];
    
    // 过滤、排序并转换
    return allSessions
      .filter(s => s.messages.length > 0) // 只显示有消息的会话
      .sort((a, b) => {
        // 按时间戳倒序排列
        return (b.timestamp || '').localeCompare(a.timestamp || '');
      })
      .map(s => {
        const lastMessage = s.messages[s.messages.length - 1];
        const preview = lastMessage?.content.slice(0, 50) || '新对话';
        
        // 格式化时间戳
        const timestamp = formatTimestamp(s.timestamp);
        
        return {
          id: s.id,
          name: s.name,
          timestamp,
          preview,
        };
      });
  }, [sessions, isHistoryOpen]);

  // 格式化时间戳为友好的显示格式
  function formatTimestamp(isoString: string): string {
    try {
      const date = new Date(isoString);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffMins = Math.floor(diffMs / 60000);
      const diffHours = Math.floor(diffMs / 3600000);
      const diffDays = Math.floor(diffMs / 86400000);
      
      if (diffMins < 1) return '刚刚';
      if (diffMins < 60) return `${diffMins} 分钟前`;
      if (diffHours < 24) return `${diffHours} 小时前`;
      if (diffDays < 7) return `${diffDays} 天前`;
      
      // 超过7天显示日期
      return date.toLocaleDateString('zh-CN', { 
        month: 'short', 
        day: 'numeric' 
      });
    } catch (error) {
      return '未知时间';
    }
  }

  // 根据第一条消息生成标题
  function generateChatTitle(firstMessage: string): string {
    // 清理消息内容：移除多余空白字符和换行
    const cleanedMessage = firstMessage.trim().replace(/\s+/g, ' ');
    
    // 截取前20个字符作为标题
    if (cleanedMessage.length <= 20) {
      return cleanedMessage;
    }
    
    return cleanedMessage.slice(0, 20) + '...';
  }

  // 检查是否需要生成标题
  function shouldGenerateTitle(session: ChatSession): boolean {
    // 必须有消息
    if (session.messages.length === 0) return false;
    
    // 获取用户消息
    const userMessages = session.messages.filter(m => m.role === 'user');
    
    // 只在有且仅有1条用户消息时生成标题
    if (userMessages.length !== 1) return false;
    
    // 标题仍为默认格式才生成
    return session.name.startsWith('Chat ');
  }

  return (
    <div className="h-screen w-full flex flex-col bg-[#ffffff] border-l border-[rgba(0,0,0,0.08)]">
      {/* Header with Tabs */}
      <div className="flex items-center justify-between px-3 py-2 border-b border-[rgba(0,0,0,0.08)] gap-2">
        {/* Tabs */}
        <div className="flex items-center gap-1 flex-1 overflow-x-auto">
          {enableTestPanel && (
            <div className="flex items-center gap-1 mr-1 shrink-0">
              <button
                type="button"
                onClick={() => setActiveView("chat")}
                className={`h-7 px-2 rounded-md text-[11px] font-medium transition-colors ${
                  activeView === "chat"
                    ? "bg-[#0a0a0a] text-[#fafafa]"
                    : "bg-[#f5f5f5] text-[#0a0a0a] hover:bg-[#e5e5e5]"
                }`}
                title="聊天"
              >
                聊天
              </button>
              <button
                type="button"
                onClick={() => setActiveView("tests")}
                className={`h-7 px-2 rounded-md text-[11px] font-medium transition-colors ${
                  activeView === "tests"
                    ? "bg-[#0a0a0a] text-[#fafafa]"
                    : "bg-[#f5f5f5] text-[#0a0a0a] hover:bg-[#e5e5e5]"
                }`}
                title="测试/诊断"
              >
                测试
              </button>
            </div>
          )}

          {activeView === "chat" ? (
            sessions.map((session) => (
              <ChatTab
                key={session.id}
                id={session.id}
                name={session.name}
                isActive={session.id === activeSessionId}
                onSelect={() => setActiveSessionId(session.id)}
                onClose={() => handleCloseChat(session.id)}
              />
            ))
          ) : (
            <div className="text-[12px] text-[#737373] px-2">
              测试/诊断（{hostType === "unknown" ? "EXCEL" : hostType.toUpperCase()}）
            </div>
          )}
        </div>

        {/* Actions - 添加右侧边距避免被 Office 插件的 i 图标遮挡 */}
        <div className="flex items-center gap-1 mr-8">
          {activeView === "chat" ? (
            <>
              <button
                onClick={handleCreateNewChat}
                aria-label="Create new chat"
                title="Create new chat"
                className="w-7 h-7 rounded-md hover:bg-[#f5f5f5] flex items-center justify-center transition-colors"
              >
                <Plus className="w-4 h-4 text-[#737373]" />
              </button>
              <button
                onClick={() => setIsHistoryOpen(true)}
                aria-label="Open history"
                title="Open history"
                className="w-7 h-7 rounded-md hover:bg-[#f5f5f5] flex items-center justify-center transition-colors"
              >
                <Clock className="w-4 h-4 text-[#737373]" />
              </button>
              <button
                onClick={() => setIsSettingsOpen(true)}
                aria-label="Open settings"
                title="Open settings"
                className="w-7 h-7 rounded-md hover:bg-[#f5f5f5] flex items-center justify-center transition-colors"
              >
                <Settings className="w-4 h-4 text-[#737373]" />
              </button>
            </>
          ) : (
            <button
              type="button"
              onClick={() => setActiveView("chat")}
              className="h-7 px-2 rounded-md bg-[#f5f5f5] hover:bg-[#e5e5e5] text-[11px] font-medium text-[#0a0a0a] transition-colors"
              title="返回聊天"
            >
              返回
            </button>
          )}
        </div>
      </div>

      {activeView === "chat" ? (
        <>
          {/* Chat Messages */}
          <div className="flex-1 overflow-y-auto px-4 py-4">
            {activeSession?.messages.length === 0 ? (
              <div className="h-full flex flex-col items-center justify-center px-4">
                <h1 className="text-[24px] font-semibold text-[#0a0a0a] mb-6">
                  DocuPilot
                </h1>
                <div className="w-full max-w-md space-y-3">
                  <p className="text-[13px] text-[#737373] mb-4 text-center">
                    Try asking me:
                  </p>
                  <button
                    onClick={() => handleSendMessage("Summarize this document", "Agent")}
                    className="w-full p-3 text-left rounded-lg bg-[#f5f5f5] hover:bg-[#e5e5e5] transition-colors border border-[rgba(0,0,0,0.06)]"
                  >
                    <p className="text-[13px] text-[#0a0a0a]">Summarize this document</p>
                  </button>
                  <button
                    onClick={() => handleSendMessage("Create a table of contents", "Agent")}
                    className="w-full p-3 text-left rounded-lg bg-[#f5f5f5] hover:bg-[#e5e5e5] transition-colors border border-[rgba(0,0,0,0.06)]"
                  >
                    <p className="text-[13px] text-[#0a0a0a]">Create a table of contents</p>
                  </button>
                  <button
                    onClick={() => handleSendMessage("Find and fix grammar errors", "Agent")}
                    className="w-full p-3 text-left rounded-lg bg-[#f5f5f5] hover:bg-[#e5e5e5] transition-colors border border-[rgba(0,0,0,0.06)]"
                  >
                    <p className="text-[13px] text-[#0a0a0a]">Find and fix grammar errors</p>
                  </button>
                </div>
              </div>
            ) : (
              <>
                {activeSession?.messages.map((message) => (
                  <MessageBubble key={message.id} message={message} />
                ))}
                <div ref={messagesEndRef} />
              </>
            )}
          </div>

          {/* Input */}
          <ChatInput onSend={handleSendMessage} onAbort={abortChatMessage} isLoading={isLoading} sessionId={currentSessionId || undefined} />
        </>
      ) : (
        <div className="flex-1 min-h-0">
          <OfficeTestPanel hostType={normalizedTestHostType} />
        </div>
      )}

      {/* History Panel */}
      <HistoryPanel
        isOpen={isHistoryOpen}
        onClose={() => setIsHistoryOpen(false)}
        onSelectHistory={(id) => handleSelectHistory(id)}
        onDeleteHistory={handleDeleteHistory}
        historyItems={historyItems}
      />

      {/* Settings Dialog */}
      <SettingsDialog
        open={isSettingsOpen}
        onOpenChange={setIsSettingsOpen}
      />
    </div>
  );
}
