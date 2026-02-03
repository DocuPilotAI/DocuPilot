"use client";

import { useEffect, useState } from "react";
import { X, ChevronUp, ChevronDown } from "lucide-react";
import { TodoItem, TaskStatus } from "@/types/chat";

interface TodoPanelProps {
  todos: TodoItem[];
  isVisible: boolean;
  onClose?: () => void;
}

const statusIcons: Record<TaskStatus, string> = {
  pending: "○",
  in_progress: "◐",
  completed: "✓",
  failed: "✗",
};

const statusColors: Record<TaskStatus, string> = {
  pending: "text-gray-400",
  in_progress: "text-blue-500",
  completed: "text-green-500",
  failed: "text-red-500",
};

export function TodoPanel({ todos, isVisible, onClose }: TodoPanelProps) {
  const [isCollapsed, setIsCollapsed] = useState(false);

  // 自动隐藏：当所有任务完成后 2 秒自动隐藏
  useEffect(() => {
    if (todos.length > 0) {
      const allCompleted = todos.every(
        (todo) => todo.status === "completed" || todo.status === "failed"
      );
      if (allCompleted && isVisible) {
        const timer = setTimeout(() => {
          onClose?.();
        }, 2000);
        return () => clearTimeout(timer);
      }
    }
  }, [todos, isVisible, onClose]);

  if (!isVisible || todos.length === 0) {
    return null;
  }

  const completedCount = todos.filter((todo) => todo.status === "completed").length;
  const totalCount = todos.length;
  const progressPercent = totalCount > 0 ? (completedCount / totalCount) * 100 : 0;

  return (
    <div className="fixed bottom-20 right-5 w-72 bg-white/95 rounded-xl shadow-lg border backdrop-blur-sm z-50 slide-in-up">
      {/* 顶部标题栏 */}
      <div 
        className="flex justify-between items-center p-3 border-b bg-blue-50/50 cursor-pointer select-none"
        onClick={() => setIsCollapsed(!isCollapsed)}
      >
        <div className="flex items-center gap-2">
          <span>📋</span>
          <span className="text-sm font-semibold">任务进度</span>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold text-blue-600 bg-blue-100 px-2 py-0.5 rounded-full">
            {completedCount}/{totalCount}
          </span>
          <button
            onClick={(e) => {
              e.stopPropagation();
              setIsCollapsed(!isCollapsed);
            }}
            className="text-gray-500 hover:text-gray-700 p-0.5"
          >
            {isCollapsed ? <ChevronDown className="w-4 h-4" /> : <ChevronUp className="w-4 h-4" />}
          </button>
          {onClose && (
            <button
              onClick={(e) => {
                e.stopPropagation();
                onClose();
              }}
              className="text-gray-400 hover:text-gray-600 p-0.5"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* 进度条 */}
      {!isCollapsed && (
        <div className="h-1 bg-gray-100">
          <div 
            className="h-full bg-blue-500 transition-all duration-300"
            style={{ width: `${progressPercent}%` }}
          />
        </div>
      )}

      {/* 任务列表 */}
      {!isCollapsed && (
        <div className="p-2 max-h-72 overflow-y-auto">
          {todos.map((todo) => (
            <div 
              key={todo.id} 
              className="flex items-start gap-2 p-2 rounded hover:bg-gray-50 transition-colors"
            >
              <span 
                className={`font-bold text-sm mt-0.5 flex-shrink-0 ${statusColors[todo.status]} ${
                  todo.status === "in_progress" ? "pulse-icon" : ""
                }`}
              >
                {statusIcons[todo.status]}
              </span>
              <span 
                className={`text-sm leading-snug ${
                  todo.status === "completed" ? "line-through text-gray-400" : "text-gray-700"
                }`}
                title={todo.content}
              >
                {todo.content}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
