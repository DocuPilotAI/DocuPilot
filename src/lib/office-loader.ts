/**
 * Office.js 动态加载器
 * 只在需要时加载 Office.js，避免在普通浏览器环境中产生警告
 */

let isOfficeJsLoaded = false;
let isOfficeJsLoading = false;
const loadCallbacks: Array<() => void> = [];

function hasForceOfficeParam(): boolean {
  if (typeof window === 'undefined') return false;
  try {
    const urlParams = new URLSearchParams(window.location.search);
    // 允许通过 query param 强制把页面当作 Office 环境
    return (
      urlParams.get('forceOffice') === '1' ||
      urlParams.get('forceOffice') === 'true' ||
      urlParams.get('isOfficeAddin') === '1' ||
      urlParams.get('isOfficeAddin') === 'true'
    );
  } catch {
    return false;
  }
}

/**
 * 检测是否在 Office Add-in 环境中
 */
export function isInOfficeEnvironment(): boolean {
  if (typeof window === 'undefined') return false;

  // 显式 override：用于测试页或调试场景
  if (hasForceOfficeParam()) return true;
  
  // 检测是否在 iframe 中（Office Add-in 通常在 iframe 中运行）
  let inIframe = false;
  try {
    inIframe = window.self !== window.top;
  } catch {
    // 跨域访问 parent/top 抛错时，认为在 iframe 中
    inIframe = true;
  }
  
  // 检测是否已经加载了 Office.js
  const hasOffice = typeof Office !== 'undefined';
  
  // 检测 URL 参数
  const urlParams = new URLSearchParams(window.location.search);
  const hasOfficeParam = urlParams.has('_host_Info') || urlParams.has('isOfficeAddin');
  
  // 兜底：部分宿主（尤其某些 Excel 环境）不一定满足 iframe/参数条件，但 userAgent 往往包含 Office 字样
  const ua = (navigator?.userAgent || '').toLowerCase();
  const looksLikeOfficeWebView =
    ua.includes('microsoft office') || ua.includes('office') || ua.includes('excel') || ua.includes('word') || ua.includes('powerpoint');

  return inIframe || hasOffice || hasOfficeParam || window.name === 'Office Add-in' || looksLikeOfficeWebView;
}

/**
 * 动态加载 Office.js
 */
export function loadOfficeJs(): Promise<void> {
  if (typeof window === 'undefined') {
    return Promise.resolve();
  }

  // 如果已经加载，直接返回
  if (isOfficeJsLoaded || typeof Office !== 'undefined') {
    isOfficeJsLoaded = true;
    return Promise.resolve();
  }

  // 如果正在加载，等待加载完成
  if (isOfficeJsLoading) {
    return new Promise((resolve) => {
      loadCallbacks.push(resolve);
    });
  }

  // 开始加载
  isOfficeJsLoading = true;

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
    script.async = true;

    script.onload = () => {
      isOfficeJsLoaded = true;
      isOfficeJsLoading = false;
      
      // 调用所有等待的回调
      loadCallbacks.forEach(callback => callback());
      loadCallbacks.length = 0;
      
      console.log('[OfficeLoader] Office.js loaded successfully');
      resolve();
    };

    script.onerror = () => {
      isOfficeJsLoading = false;
      console.error('[OfficeLoader] Failed to load Office.js');
      reject(new Error('Failed to load Office.js'));
    };

    document.head.appendChild(script);
  });
}

/**
 * 等待 Office 准备就绪
 */
export function waitForOfficeReady(options?: { timeoutMs?: number }): Promise<Office.HostType | null> {
  if (typeof window === 'undefined') {
    return Promise.resolve(null);
  }

  // 如果不在 Office 环境中，直接返回
  if (!isInOfficeEnvironment()) {
    return Promise.resolve(null);
  }

  const timeoutMs = options?.timeoutMs ?? 10000;

  return new Promise((resolve) => {
    let settled = false;

    const done = (host: Office.HostType | null) => {
      if (settled) return;
      settled = true;
      resolve(host);
    };

    const timer = window.setTimeout(() => {
      console.warn('[OfficeLoader] Office.onReady timeout');
      done(null);
    }, timeoutMs);

    const onReady = () => {
      try {
        if (typeof Office !== 'undefined' && Office.onReady) {
          Office.onReady((info) => {
            window.clearTimeout(timer);
            console.log('[OfficeLoader] Office ready:', info.host);
            done(info.host);
          });
        } else {
          window.clearTimeout(timer);
          done(null);
        }
      } catch {
        window.clearTimeout(timer);
        done(null);
      }
    };

    if (typeof Office !== 'undefined' && (Office as any).onReady) {
      onReady();
      return;
    }

    // 加载 Office.js 后再等待就绪
    loadOfficeJs()
      .then(() => onReady())
      .catch(() => {
        window.clearTimeout(timer);
        done(null);
      });
  });
}
