"use client";

import { motion } from 'framer-motion';
import { ChatMessage } from "@/types/chat";
import { ToolCallDisplay } from "./ToolCallDisplay";
import { MarkdownRenderer } from "./MarkdownRenderer";
import { TaskListMessage } from "./TaskListMessage";

interface MessageBubbleProps {
  message: ChatMessage;
}

export function MessageBubble({ message }: MessageBubbleProps) {
  // 任务列表消息
  if (message.role === "task_list") {
    return (
      <TaskListMessage
        tasks={message.metadata?.tasks || []}
        title={message.metadata?.taskTitle}
        objective={message.metadata?.taskObjective}
        timestamp={message.timestamp}
      />
    );
  }

  // 用户消息
  if (message.role === "user") {
    return (
      <motion.div
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.2 }}
        className="flex justify-end mb-4"
      >
        <div className="max-w-[85%] items-end flex flex-col gap-1">
          <div className="px-3 py-2.5 rounded-lg bg-[#0a0a0a] text-[#fafafa]">
            <p className="text-[13px] leading-relaxed tracking-[-0.01em] whitespace-pre-wrap">
              {message.content}
            </p>
          </div>
          {message.timestamp && (
            <span className="text-[10px] text-[#a3a3a3] px-1">
              {message.timestamp}
            </span>
          )}
        </div>
      </motion.div>
    );
  }

  // 助手消息（包括assistant, error等其他角色）
  const isStreamingPlaceholder = message.streaming && message.content.trim().length === 0;
  
  return (
    <motion.div
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.2 }}
      className="flex justify-start mb-4"
    >
      <div className="max-w-[85%] items-start flex flex-col gap-1">
        <div className="px-3 py-2.5 rounded-lg bg-[#f5f5f5] border border-[rgba(0,0,0,0.08)]">
          {isStreamingPlaceholder ? (
            <div className="thinking-dots" aria-label="正在思考中">
              <span />
              <span />
              <span />
            </div>
          ) : (
            <>
              <div className="text-[13px] leading-relaxed tracking-[-0.01em] text-[#0a0a0a]">
                <MarkdownRenderer content={message.content} />
              </div>

              {/* 工具调用显示 */}
              {message.metadata?.toolName && (
                <div className="mt-2">
                  <ToolCallDisplay
                    toolName={message.metadata.toolName}
                    input={message.metadata.toolInput}
                    result={message.metadata.toolResult}
                  />
                </div>
              )}

              {/* 重试状态显示 */}
              {message.metadata?.retryCount && (
                <div className="mt-2 pt-2 border-t border-[rgba(0,0,0,0.08)]">
                  <div className="flex items-center gap-2 text-[11px] text-[#737373]">
                    <span className="inline-block">🔄</span>
                    <span>
                      {message.metadata.finalError 
                        ? `已尝试 ${message.metadata.retryCount} 次`
                        : `重试 ${message.metadata.retryCount}/${message.metadata.maxRetries}`
                      }
                    </span>
                    {message.metadata.errorType && (
                      <span className="px-1.5 py-0.5 rounded bg-[#fef2f2] text-[#dc2626] text-[10px]">
                        {message.metadata.errorType}
                      </span>
                    )}
                  </div>
                </div>
              )}
            </>
          )}
        </div>
        {message.timestamp && (
          <span className="text-[10px] text-[#a3a3a3] px-1">
            {message.timestamp}
          </span>
        )}
      </div>
    </motion.div>
  );
}
