"use client";

import { useEffect } from 'react';

export default function Home() {
  useEffect(() => {
    // 在 Office Add-in 环境中使用 window.location 而不是 Next.js 路由
    if (typeof window !== 'undefined') {
      window.location.href = '/taskpane';
    }
  }, []);

  return (
    <div className="flex items-center justify-center h-screen bg-gray-50">
      <div className="text-center">
        <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin mx-auto mb-4" />
        <p className="text-gray-600">加载中...</p>
      </div>
    </div>
  );
}
