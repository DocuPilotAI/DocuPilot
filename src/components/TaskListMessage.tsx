"use client";

import { useState, useEffect, useMemo } from "react";
import { motion, AnimatePresence } from "framer-motion";
import { ChevronDown, ChevronUp } from "lucide-react";
import { TodoItem, TaskStatus } from "@/types/chat";
import { cn } from "@/lib/utils";

interface TaskListMessageProps {
  tasks: TodoItem[];
  title?: string;
  objective?: string;
  timestamp?: string;
}

const statusIcons: Record<TaskStatus, string> = {
  pending: "○",
  in_progress: "◐",
  completed: "✓",
  failed: "✗",
};

const statusColors: Record<TaskStatus, string> = {
  pending: "text-[#a3a3a3]",
  in_progress: "text-[#0a0a0a]",
  completed: "text-[#10b981]",
  failed: "text-[#ef4444]",
};

export function TaskListMessage({ 
  tasks, 
  title = "任务规划", 
  objective,
  timestamp 
}: TaskListMessageProps) {
  const [isExpanded, setIsExpanded] = useState(true);
  const [isVisible, setIsVisible] = useState(true);

  // 计算进度
  const progress = useMemo(() => {
    const completed = tasks.filter(t => t.status === "completed").length;
    const total = tasks.length;
    return { 
      completed, 
      total, 
      percentage: total > 0 ? (completed / total) * 100 : 0 
    };
  }, [tasks]);

  // 获取当前执行的任务
  const currentTask = useMemo(() => {
    return tasks.find(t => t.status === "in_progress");
  }, [tasks]);

  // 所有任务完成后延迟淡出（可选）
  useEffect(() => {
    if (progress.completed === progress.total && progress.total > 0) {
      // 注释掉自动隐藏功能，保留在历史记录中
      // const timer = setTimeout(() => {
      //   setIsVisible(false);
      // }, 2000);
      // return () => clearTimeout(timer);
    }
  }, [progress]);

  if (!isVisible) {
    return null;
  }

  return (
    <motion.div
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0 }}
      transition={{ duration: 0.3 }}
      className="flex justify-start mb-4"
    >
      <div className="max-w-[85%] items-start flex flex-col gap-1">
        <div className="w-full px-4 py-3 rounded-lg bg-[#f5f5f5] border border-[rgba(0,0,0,0.08)] shadow-sm">
          {/* 任务头部 */}
          <div 
            className="flex items-center justify-between cursor-pointer select-none mb-2"
            onClick={() => setIsExpanded(!isExpanded)}
          >
            <div className="flex items-center gap-2 flex-1">
              <span className="text-sm">📋</span>
              <span className="text-sm font-semibold text-[#0a0a0a]">
                {title}
              </span>
              <span className="bg-[#0a0a0a] text-white text-[10px] px-2 py-0.5 rounded-full font-semibold">
                {progress.completed}/{progress.total}
              </span>
            </div>
            <button className="text-[#a3a3a3] hover:text-[#0a0a0a] transition-colors p-1">
              {isExpanded ? (
                <ChevronUp className="w-4 h-4" />
              ) : (
                <ChevronDown className="w-4 h-4" />
              )}
            </button>
          </div>

          {/* 目标描述 */}
          {objective && (
            <div className="text-xs text-[#737373] mb-3 leading-relaxed">
              {objective}
            </div>
          )}

          {/* 当前任务（折叠时显示） */}
          {!isExpanded && currentTask && (
            <div className="flex items-start gap-2 py-1">
              <span className={cn(
                "font-bold text-sm mt-0.5 flex-shrink-0 animate-pulse",
                statusColors[currentTask.status]
              )}>
                {statusIcons[currentTask.status]}
              </span>
              <span className="text-[13px] text-[#0a0a0a] leading-relaxed">
                正在执行：{currentTask.content}
              </span>
            </div>
          )}

          {/* 任务列表（展开时显示） */}
          <AnimatePresence>
            {isExpanded && (
              <motion.div
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: "auto", opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                transition={{ duration: 0.3, ease: "easeInOut" }}
                className="overflow-hidden"
              >
                <div className="space-y-1 mt-2">
                  {tasks.map((task, index) => {
                    const isCompleted = task.status === "completed";
                    const isFailed = task.status === "failed";
                    // 确保使用唯一的 key，如果 task.id 不存在则使用 index
                    const taskKey = task.id || `task-${index}`;
                    
                    return (
                      <motion.div
                        key={taskKey}
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: index * 0.05 }}
                        className="flex items-start gap-2 py-1.5 px-2 hover:bg-white/50 rounded transition-colors"
                      >
                        <span 
                          className={cn(
                            "font-bold text-sm mt-0.5 flex-shrink-0",
                            statusColors[task.status],
                            task.status === "in_progress" && "animate-pulse"
                          )}
                        >
                          {statusIcons[task.status]}
                        </span>
                        <div className="flex-1 min-w-0">
                          <span 
                            className={cn(
                              "text-[13px] leading-relaxed transition-all duration-200",
                              isCompleted && "line-through text-[#a3a3a3]",
                              isFailed && "line-through text-[#ef4444]",
                              !isCompleted && !isFailed && "text-[#0a0a0a]"
                            )}
                          >
                            <span className="font-semibold text-[#737373] mr-1">
                              {index + 1}.
                            </span>
                            {task.content}
                          </span>
                          {/* 显示错误信息 */}
                          {task.error && (
                            <div className="text-xs text-[#ef4444] mt-1">
                              {task.error}
                            </div>
                          )}
                        </div>
                      </motion.div>
                    );
                  })}
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          {/* 进度条 */}
          <div className="mt-3 h-1 bg-[rgba(0,0,0,0.05)] rounded-full overflow-hidden">
            <motion.div 
              className="h-full bg-[#0a0a0a] transition-all duration-300"
              initial={{ width: 0 }}
              animate={{ width: `${progress.percentage}%` }}
              transition={{ duration: 0.5, ease: "easeOut" }}
            />
          </div>
        </div>
        
        {/* 时间戳 */}
        {timestamp && (
          <span className="text-[10px] text-[#a3a3a3] px-1">
            {timestamp}
          </span>
        )}
      </div>
    </motion.div>
  );
}
