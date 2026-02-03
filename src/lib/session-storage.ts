import { ChatMessage } from "@/types/chat";

const STORAGE_KEY = "docupilot:sessions";
const LEGACY_SESSION_KEY = "docupilot_session_id";

export interface StoredSession {
  id: string;
  name: string;
  sessionId?: string; // 后端的 session ID
  messages: ChatMessage[];
  timestamp: string; // ISO 格式的时间戳
}

/**
 * 从 localStorage 加载所有会话
 */
export function loadSessions(): StoredSession[] {
  if (typeof window === "undefined") {
    return [];
  }
  
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return [];
    }
    
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      console.warn("[session-storage] Invalid data format, expected array");
      return [];
    }
    
    return parsed;
  } catch (error) {
    console.error("[session-storage] Failed to load sessions:", error);
    return [];
  }
}

/**
 * 保存所有会话到 localStorage
 */
export function saveSessions(sessions: StoredSession[]): boolean {
  if (typeof window === "undefined") {
    return false;
  }
  
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    return true;
  } catch (error) {
    console.error("[session-storage] Failed to save sessions:", error);
    return false;
  }
}

/**
 * 删除指定会话
 */
export function deleteSession(sessionId: string): boolean {
  const sessions = loadSessions();
  const filtered = sessions.filter((s) => s.id !== sessionId);
  return saveSessions(filtered);
}

/**
 * 更新或添加单个会话
 */
export function upsertSession(session: StoredSession): boolean {
  const sessions = loadSessions();
  const existingIndex = sessions.findIndex((s) => s.id === session.id);
  
  if (existingIndex >= 0) {
    sessions[existingIndex] = session;
  } else {
    sessions.push(session);
  }
  
  return saveSessions(sessions);
}

/**
 * 获取单个会话
 */
export function getSession(sessionId: string): StoredSession | null {
  const sessions = loadSessions();
  return sessions.find((s) => s.id === sessionId) || null;
}

/**
 * 迁移旧版本的 sessionId（如果存在）
 * 这是为了向后兼容
 */
export function migrateLegacySession(): string | null {
  if (typeof window === "undefined") {
    return null;
  }
  
  try {
    const legacySessionId = localStorage.getItem(LEGACY_SESSION_KEY);
    if (legacySessionId) {
      console.log("[session-storage] Found legacy session ID:", legacySessionId);
      // 不删除旧的 key，让 useChat 继续使用
      return legacySessionId;
    }
  } catch (error) {
    console.error("[session-storage] Failed to migrate legacy session:", error);
  }
  
  return null;
}

/**
 * 清除所有会话数据（用于测试或重置）
 */
export function clearAllSessions(): boolean {
  if (typeof window === "undefined") {
    return false;
  }
  
  try {
    localStorage.removeItem(STORAGE_KEY);
    return true;
  } catch (error) {
    console.error("[session-storage] Failed to clear sessions:", error);
    return false;
  }
}
