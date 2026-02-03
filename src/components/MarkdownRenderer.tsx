"use client";

import React, { useEffect, useRef } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeRaw from 'rehype-raw';
import mermaid from 'mermaid';

interface MarkdownRendererProps {
  content: string;
}

// 初始化 mermaid
mermaid.initialize({ startOnLoad: true, theme: 'default' });

export function MarkdownRenderer({ content }: MarkdownRendererProps) {
  const mermaidRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    // 处理 mermaid 图表
    if (mermaidRef.current) {
      const mermaidElements = mermaidRef.current.querySelectorAll('.mermaid');
      mermaidElements.forEach((element) => {
        try {
          mermaid.contentLoaded();
        } catch (error) {
          console.error('Mermaid rendering error:', error);
        }
      });
    }
  }, [content]);

  return (
    <div
      ref={mermaidRef}
      className="markdown-content prose prose-sm max-w-none"
    >
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[rehypeRaw]}
        components={{
          // 自定义代码块处理
          code({ node, inline, className, children, ...props }: any) {
            const match = /language-(\w+)/.exec(className || '');
            const language = match ? match[1] : '';

            // 检查是否是 mermaid 图表
            if (language === 'mermaid') {
              return (
                <div className="mermaid">
                  {String(children).replace(/\n$/, '')}
                </div>
              );
            }

            // 普通代码块
            if (!inline && language) {
              return (
                <pre className="bg-[#f5f5f5] border border-[rgba(0,0,0,0.08)] rounded-lg p-3 overflow-x-auto">
                  <code className={className} {...props}>
                    {children}
                  </code>
                </pre>
              );
            }

            // 行内代码
            return (
              <code
                className="bg-[#f5f5f5] px-1.5 py-0.5 rounded text-[#d32f2f] font-mono text-[12px]"
                {...props}
              >
                {children}
              </code>
            );
          },

          // 自定义标题
          h1: ({ children }) => (
            <h1 className="text-[20px] font-bold mt-4 mb-2 text-[#0a0a0a]">
              {children}
            </h1>
          ),
          h2: ({ children }) => (
            <h2 className="text-[18px] font-bold mt-3 mb-2 text-[#0a0a0a]">
              {children}
            </h2>
          ),
          h3: ({ children }) => (
            <h3 className="text-[16px] font-bold mt-2 mb-1 text-[#0a0a0a]">
              {children}
            </h3>
          ),
          h4: ({ children }) => (
            <h4 className="text-[14px] font-bold mt-2 mb-1 text-[#0a0a0a]">
              {children}
            </h4>
          ),
          h5: ({ children }) => (
            <h5 className="text-[13px] font-bold mt-1 mb-1 text-[#0a0a0a]">
              {children}
            </h5>
          ),
          h6: ({ children }) => (
            <h6 className="text-[12px] font-bold mt-1 mb-1 text-[#0a0a0a]">
              {children}
            </h6>
          ),

          // 自定义列表
          ul: ({ children }) => (
            <ul className="list-disc list-inside ml-2 my-2 space-y-1">
              {children}
            </ul>
          ),
          ol: ({ children }) => (
            <ol className="list-decimal list-inside ml-2 my-2 space-y-1">
              {children}
            </ol>
          ),
          li: ({ children }) => (
            <li className="text-[13px] text-[#0a0a0a]">
              {children}
            </li>
          ),

          // 自定义表格
          table: ({ children }) => (
            <table className="border-collapse border border-[rgba(0,0,0,0.1)] my-2 w-full">
              {children}
            </table>
          ),
          thead: ({ children }) => (
            <thead className="bg-[#f5f5f5]">
              {children}
            </thead>
          ),
          tbody: ({ children }) => (
            <tbody>
              {children}
            </tbody>
          ),
          tr: ({ children }) => (
            <tr className="border border-[rgba(0,0,0,0.1)]">
              {children}
            </tr>
          ),
          th: ({ children }) => (
            <th className="border border-[rgba(0,0,0,0.1)] px-2 py-1 text-left font-bold text-[12px]">
              {children}
            </th>
          ),
          td: ({ children }) => (
            <td className="border border-[rgba(0,0,0,0.1)] px-2 py-1 text-[12px]">
              {children}
            </td>
          ),

          // 自定义段落
          p: ({ children }) => (
            <p className="text-[13px] leading-relaxed text-[#0a0a0a] my-1">
              {children}
            </p>
          ),

          // 自定义链接
          a: ({ href, children }) => (
            <a
              href={href}
              target="_blank"
              rel="noopener noreferrer"
              className="text-[#0066cc] hover:underline"
            >
              {children}
            </a>
          ),

          // 自定义强调
          strong: ({ children }) => (
            <strong className="font-bold text-[#0a0a0a]">
              {children}
            </strong>
          ),
          em: ({ children }) => (
            <em className="italic text-[#0a0a0a]">
              {children}
            </em>
          ),

          // 自定义引用块
          blockquote: ({ children }) => (
            <blockquote className="border-l-4 border-[#0066cc] pl-3 my-2 text-[#666666] italic">
              {children}
            </blockquote>
          ),

          // 自定义水平线
          hr: () => (
            <hr className="my-2 border-t border-[rgba(0,0,0,0.1)]" />
          ),
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  );
}
