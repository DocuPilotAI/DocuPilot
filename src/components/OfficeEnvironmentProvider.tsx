"use client";

import { useEffect, useState } from 'react';
// 导入 polyfills 以确保在客户端最早阶段执行
import '@/lib/office-polyfills';

/**
 * Office Add-in 环境兼容性处理
 * 在 Office WebView 中，某些浏览器 API 可能不可用或受限
 */
export function OfficeEnvironmentProvider({ children }: { children: React.ReactNode }) {
  const [isReady, setIsReady] = useState(false);

  useEffect(() => {
    // 检测是否在 Office Add-in 环境中
    // 注意：Office 对象可能由 office-loader 异步加载，所以这里不仅仅依赖 typeof Office
    console.log('[OfficeEnvironment] Provider mounted');
    
    // 再次检查 polyfills 是否生效（双重保险）
    if (typeof window !== 'undefined' && window.history) {
      if (!window.history.replaceState) {
        console.warn('[OfficeEnvironment] History API missing after polyfills, reapplying...');
        // 强制应用简单的 polyfill
        window.history.replaceState = () => {};
        window.history.pushState = () => {};
      }
    }

    // 禁用某些可能导致问题的 Next.js 功能
    if (typeof window !== 'undefined') {
      // 阻止 Next.js 的自动预取
      (window as any).__NEXT_DATA__ = (window as any).__NEXT_DATA__ || {};
      (window as any).__NEXT_DATA__.autoExport = true;
    }
    
    // 标记为准备就绪
    setIsReady(true);
  }, []);

  // 避免 hydration mismatch，在客户端准备好之前不渲染子组件
  if (!isReady) {
    return null;
  }

  return <>{children}</>;
}
