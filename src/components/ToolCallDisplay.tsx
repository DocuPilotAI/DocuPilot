"use client";

import { useState } from "react";
import { ChevronDown, ChevronUp, Wrench } from "lucide-react";

interface ToolCallDisplayProps {
  toolName: string;
  input?: unknown;
  result?: unknown;
}

export function ToolCallDisplay({ toolName, input, result }: ToolCallDisplayProps) {
  const [isExpanded, setIsExpanded] = useState(false);
  
  // 格式化工具名称
  const displayName = toolName
    .replace("office_excel_", "Excel: ")
    .replace("office_word_", "Word: ")
    .replace("office_ppt_", "PPT: ")
    .replace(/_/g, " ");
  
  return (
    <div className="mt-2 border border-gray-200 rounded-lg overflow-hidden">
      {/* 工具名称 */}
      <button
        onClick={() => setIsExpanded(!isExpanded)}
        className="w-full flex items-center justify-between p-2 bg-gray-50 hover:bg-gray-100 transition-colors"
      >
        <div className="flex items-center gap-2">
          <Wrench className="w-4 h-4 text-gray-500" />
          <span className="text-sm font-medium">{displayName}</span>
        </div>
        {isExpanded ? (
          <ChevronUp className="w-4 h-4 text-gray-500" />
        ) : (
          <ChevronDown className="w-4 h-4 text-gray-500" />
        )}
      </button>
      
      {/* 展开显示输入和输出 */}
      {isExpanded && (
        <div className="p-2 text-xs font-mono bg-gray-900 text-gray-100 max-h-60 overflow-auto">
          {input ? (
            <div className="mb-2">
              <div className="text-gray-400 mb-1">输入:</div>
              <pre className="whitespace-pre-wrap break-all">
                {JSON.stringify(input, null, 2)}
              </pre>
            </div>
          ) : null}
          {result ? (
            <div>
              <div className="text-gray-400 mb-1">结果:</div>
              <pre className="whitespace-pre-wrap break-all">
                {JSON.stringify(result, null, 2)}
              </pre>
            </div>
          ) : null}
        </div>
      )}
    </div>
  );
}
