"use client";

/**
 * Office Add-in 环境的 Polyfills
 * Office 任务窗格运行在受限的浏览器环境中，某些 Web API 不可用
 * 这个文件应该在应用启动的最早阶段被导入
 */

// 仅在客户端执行
if (typeof window !== 'undefined') {
  console.log('[Office Polyfills] Initializing...');

  // 1. 确保 window.history 存在
  if (!window.history) {
    console.log('[Office Polyfills] Creating window.history object');
    (window as any).history = {};
  }

  // 2. Polyfill replaceState
  if (!window.history.replaceState || typeof window.history.replaceState !== 'function') {
    console.log('[Office Polyfills] Polyfilling history.replaceState');
    try {
      window.history.replaceState = function(data: any, title: string, url?: string | URL | null) {
        console.debug('[Office Polyfills] history.replaceState called (no-op):', url);
      };
    } catch (e) {
      console.warn('[Office Polyfills] Failed to assign replaceState directly, trying defineProperty', e);
      try {
        Object.defineProperty(window.history, 'replaceState', {
          value: function(data: any, title: string, url?: string | URL | null) {
            console.debug('[Office Polyfills] history.replaceState called (no-op via defineProperty):', url);
          },
          writable: true,
          configurable: true
        });
      } catch (e2) {
        console.error('[Office Polyfills] Failed to polyfill replaceState', e2);
      }
    }
  }

  // 3. Polyfill pushState
  if (!window.history.pushState || typeof window.history.pushState !== 'function') {
    console.log('[Office Polyfills] Polyfilling history.pushState');
    try {
      window.history.pushState = function(data: any, title: string, url?: string | URL | null) {
        console.debug('[Office Polyfills] history.pushState called (no-op):', url);
      };
    } catch (e) {
      console.warn('[Office Polyfills] Failed to assign pushState directly, trying defineProperty', e);
      try {
        Object.defineProperty(window.history, 'pushState', {
          value: function(data: any, title: string, url?: string | URL | null) {
            console.debug('[Office Polyfills] history.pushState called (no-op via defineProperty):', url);
          },
          writable: true,
          configurable: true
        });
      } catch (e2) {
        console.error('[Office Polyfills] Failed to polyfill pushState', e2);
      }
    }
  }

  // 4. 确保 history.state 可访问
  // 注意：这里我们检查属性描述符，或者尝试访问它
  try {
    const stateDesc = Object.getOwnPropertyDescriptor(window.history, 'state');
    if (!stateDesc || (!stateDesc.get && !stateDesc.value)) {
      console.log('[Office Polyfills] Polyfilling history.state');
      Object.defineProperty(window.history, 'state', {
        get: function() {
          return {};
        },
        configurable: true
      });
    }
  } catch (e) {
    console.warn('[Office Polyfills] Error checking/defining history.state:', e);
  }
  
  // 5. 其他 History 方法
  ['back', 'forward', 'go'].forEach(method => {
    if (!(window.history as any)[method]) {
      (window.history as any)[method] = () => {
        console.debug(`[Office Polyfills] history.${method} called (no-op)`);
      };
    }
  });

  console.log('[Office Polyfills] Polyfills applied successfully');
}

export {};
