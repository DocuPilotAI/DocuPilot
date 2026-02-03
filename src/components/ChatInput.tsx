"use client";

import { useState, useRef, useEffect, KeyboardEvent } from "react";
// import Link from "next/link"; // Link 组件在 Office 环境中可能导致 history API 错误
import { Send, ChevronDown, Paperclip, X, Square } from "lucide-react";

interface ChatInputProps {
  onSend: (message: string, mode: 'Agent' | 'Plan', uploadedFiles?: UploadedFile[]) => void;
  onAbort?: () => void;
  isLoading?: boolean;
  placeholder?: string;
  disabled?: boolean;
  sessionId?: string; // 当前会话 ID
}

interface UploadedFile {
  name: string;
  size: number;
  file: File;
  serverPath?: string;  // 服务器返回的相对路径
  fileId?: string;      // 文件唯一标识
}

export function ChatInput({ 
  onSend, 
  onAbort,
  isLoading = false, 
  placeholder = "Ask DocuPilot...",
  disabled = false,
  sessionId
}: ChatInputProps) {
  const [message, setMessage] = useState("");
  const [mode, setMode] = useState<'Agent' | 'Plan'>('Agent');
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const adjustHeight = () => {
    const textarea = textareaRef.current;
    if (textarea) {
      textarea.style.height = 'auto';
      const lineHeight = 1.5 * 13; // 1.5 line-height * 13px font-size
      const maxHeight = lineHeight * 8 + 16; // 8 lines + padding (8px top + 8px bottom)
      const newHeight = Math.min(textarea.scrollHeight, maxHeight);
      textarea.style.height = `${newHeight}px`;
    }
  };

  useEffect(() => {
    adjustHeight();
  }, [message]);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (message.trim() && !isLoading && !disabled) {
      // 将上传的文件信息传递给 onSend
      onSend(message.trim(), mode, uploadedFiles.length > 0 ? uploadedFiles : undefined);
      setMessage("");
      // 清空已上传的文件
      setUploadedFiles([]);
    }
  };

  const handleKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSubmit(e as any);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setMessage(e.target.value);
  };

  const handleFileUpload = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      const newFiles: UploadedFile[] = Array.from(files).map((file) => ({
        name: file.name,
        size: file.size,
        file,
      }));
      
      setUploadedFiles((prev) => [...prev, ...newFiles]);
      
      // 上传文件到服务器并更新路径信息
      for (let i = 0; i < newFiles.length; i++) {
        const fileInfo = newFiles[i];
        const result = await uploadFile(fileInfo.file);
        if (result) {
          // 更新文件列表中的服务器路径信息
          setUploadedFiles((prev) => 
            prev.map((f) => 
              f.file === fileInfo.file 
                ? { ...f, serverPath: result.path, fileId: result.fileId }
                : f
            )
          );
        }
      }
    }
    
    // 清空 input 以允许重复上传同一文件
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  const uploadFile = async (file: File): Promise<{ path: string; fileId: string } | null> => {
    try {
      const formData = new FormData();
      formData.append("file", file);
      if (sessionId) {
        formData.append("sessionId", sessionId);
      }

      const response = await fetch("/api/files/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error(`Upload failed: ${response.statusText}`);
      }

      const result = await response.json();
      console.log("[ChatInput] File uploaded:", result);
      return { path: result.path, fileId: result.fileId };
    } catch (error) {
      console.error("[ChatInput] Upload error:", error);
      alert(`文件上传失败: ${(error as Error).message}`);
      return null;
    }
  };

  const removeFile = (index: number) => {
    setUploadedFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(1) + " MB";
  };

  return (
    <form onSubmit={handleSubmit} className="border-t border-[rgba(0,0,0,0.08)] bg-[#ffffff] p-3">
      <div className="flex flex-col gap-2">
        {/* 文件上传input */}
        <input
          ref={fileInputRef}
          type="file"
          multiple
          className="hidden"
          onChange={handleFileChange}
          accept=".txt,.pdf,.doc,.docx,.xls,.xlsx,.csv,.json,.png,.jpg,.jpeg"
        />
        
        {/* 文件预览区域 */}
        {uploadedFiles.length > 0 && (
          <div className="flex flex-wrap gap-2 p-2 bg-[#f5f5f5] rounded-lg border border-[rgba(0,0,0,0.08)]">
            {uploadedFiles.map((fileInfo, index) => (
              <div
                key={index}
                className="flex items-center gap-2 px-3 py-2 bg-white rounded-md border border-[rgba(0,0,0,0.08)] text-xs"
              >
                <span className="text-[#0a0a0a] truncate max-w-[150px]">
                  {fileInfo.name}
                </span>
                <span className="text-[#737373]">
                  {formatFileSize(fileInfo.size)}
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(index)}
                  className="w-4 h-4 flex items-center justify-center rounded-full hover:bg-[#e5e5e5] transition-colors"
                  aria-label="删除文件"
                >
                  <X className="w-3 h-3 text-[#737373]" />
                </button>
              </div>
            ))}
          </div>
        )}
        
        <div className="relative">
          <textarea
            ref={textareaRef}
            value={message}
            onChange={handleChange}
            onKeyDown={handleKeyDown}
            placeholder={placeholder}
            disabled={disabled || isLoading}
            rows={1}
            className="
              w-full min-h-[32px] px-3 py-2
              bg-[#fafafa] 
              rounded-lg 
              text-[13px] text-[#0a0a0a]
              placeholder:text-[#a3a3a3]
              focus:outline-none focus:ring-1 focus:ring-[rgba(0,0,0,0.15)]
              resize-none
              transition-all
              max-h-32
              overflow-y-auto
              disabled:opacity-50 disabled:cursor-not-allowed
            "
            style={{
              fontFamily: "'Inter', monospace",
              lineHeight: '1.5',
            }}
          />
        </div>

        <div className="flex items-center justify-between gap-2">
          <div className="relative">
            <button
              type="button"
              onClick={() => setIsDropdownOpen(!isDropdownOpen)}
              className="flex items-center justify-center gap-1 h-8 px-2.5 rounded-md hover:bg-[#f5f5f5] transition-colors"
              disabled={disabled || isLoading}
            >
              <span className="text-[11px] font-medium text-[#0a0a0a]">{mode}</span>
              <ChevronDown className="w-3 h-3 text-[#737373]" />
            </button>
            
            {isDropdownOpen && (
              <>
                <div 
                  className="fixed inset-0 z-10" 
                  onClick={() => setIsDropdownOpen(false)}
                />
                <div className="absolute bottom-full left-0 mb-1 w-24 bg-[#ffffff] border border-[rgba(0,0,0,0.08)] rounded-md shadow-lg overflow-hidden z-20">
                  <button
                    type="button"
                    onClick={() => {
                      setMode('Agent');
                      setIsDropdownOpen(false);
                    }}
                    className={`w-full px-3 py-2 text-left text-[11px] font-medium transition-colors ${
                      mode === 'Agent' 
                        ? 'bg-[#0a0a0a] text-[#fafafa]' 
                        : 'text-[#0a0a0a] hover:bg-[#f5f5f5]'
                    }`}
                  >
                    Agent
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setMode('Plan');
                      setIsDropdownOpen(false);
                    }}
                    className={`w-full px-3 py-2 text-left text-[11px] font-medium transition-colors ${
                      mode === 'Plan' 
                        ? 'bg-[#0a0a0a] text-[#fafafa]' 
                        : 'text-[#0a0a0a] hover:bg-[#f5f5f5]'
                    }`}
                  >
                    Plan
                  </button>
                </div>
              </>
            )}
          </div>

          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={handleFileUpload}
              disabled={disabled || isLoading}
              className="
                flex items-center justify-center
                w-8 h-8
                rounded-md
                hover:bg-[#f5f5f5]
                disabled:opacity-50
                disabled:cursor-not-allowed
                transition-colors
              "
              title="上传文件"
            >
              <Paperclip className="w-4 h-4 text-[#737373]" />
            </button>
            {isLoading ? (
              <button
                type="button"
                onClick={onAbort}
                aria-label="Stop generation"
                title="停止生成"
                className="
                  flex items-center justify-center 
                  w-8 h-8 
                  rounded-md 
                  bg-red-600
                  hover:bg-red-700
                  transition-colors
                "
              >
                <Square className="w-4 h-4 text-white fill-white" />
              </button>
            ) : (
              <button
                type="submit"
                disabled={!message.trim() || disabled}
                aria-label="Send message"
                title="Send message"
                className="
                  flex items-center justify-center 
                  w-8 h-8 
                  rounded-md 
                  bg-[#0a0a0a] 
                  hover:bg-[#262626]
                  disabled:bg-[#e5e5e5]
                  disabled:cursor-not-allowed
                  transition-colors
                "
              >
                <Send className={`w-4 h-4 ${message.trim() && !disabled ? 'text-[#fafafa]' : 'text-[#a3a3a3]'}`} />
              </button>
            )}
          </div>
        </div>
      </div>
    </form>
  );
}
