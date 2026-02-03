"use client";

/**
 * Office 代码执行器
 * 用于提取和执行隐藏在消息中的 Office.js 代码
 */

const DEBUG_EXECUTOR =
  process.env.NODE_ENV === "development" ||
  (typeof process !== "undefined" && process.env?.NEXT_PUBLIC_DEBUG_CHAT === "1");

export interface OfficeCodeBlock {
  host: 'word' | 'excel' | 'powerpoint';
  code: string;
}

export interface ExecutionError {
  type: 'InvalidArgument' | 'InvalidReference' | 'ApiNotFound' | 'GeneralException' | 'NetworkError' | 'UnknownError';
  code?: string;
  message: string;
  stackTrace?: string;
  debugInfo?: any;
}

export interface ErrorReport {
  timestamp: string;
  testCaseId?: string;
  hostType: 'excel' | 'word' | 'powerpoint';
  errorType: ExecutionError['type'];
  errorCode?: string;
  errorMessage: string;
  stackTrace?: string;
  userInput?: string;
  generatedCode: string;
  context: {
    officeVersion: string;
    platform: string;
    browserInfo?: string;
  };
}

/**
 * 从消息内容中提取隐藏的 Office 代码
 * 格式：<!--OFFICE-CODE:host\n代码内容\n-->
 */
export function extractHiddenOfficeCode(content: string): OfficeCodeBlock[] {
  const regex = /<!--\s*OFFICE-CODE:(word|excel|powerpoint)\s*([\s\S]*?)-->/g;
  const matches: OfficeCodeBlock[] = [];
  let match;
  
  while ((match = regex.exec(content)) !== null) {
    matches.push({
      host: match[1] as 'word' | 'excel' | 'powerpoint',
      code: match[2].trim()
    });
  }
  
  if (DEBUG_EXECUTOR && matches.length > 0) {
    console.group('[Office Code Executor] Extracted code blocks');
    matches.forEach((block, index) => {
      console.log(`Block ${index + 1}:`, {
        host: block.host,
        codeLength: block.code.length,
        preview: block.code.substring(0, 100) + '...'
      });
    });
    console.groupEnd();
  }
  
  return matches;
}

/**
 * 移除消息中的隐藏代码标记，返回用户可见的内容
 */
export function removeHiddenOfficeCode(content: string): string {
  return content.replace(/<!--\s*OFFICE-CODE:(?:word|excel|powerpoint)\s*[\s\S]*?-->/g, '');
}

/**
 * 解析 Office.js 错误类型
 */
function parseOfficeError(error: any): ExecutionError {
  // Office.js 特定错误
  if (error?.code) {
    const errorCode = String(error.code);
    
    // 根据错误代码确定类型
    if (errorCode === 'InvalidArgument') {
      return {
        type: 'InvalidArgument',
        code: errorCode,
        message: error.message || '参数无效或缺少，或格式不正确',
        debugInfo: error.debugInfo,
      };
    } else if (errorCode === 'InvalidReference') {
      return {
        type: 'InvalidReference',
        code: errorCode,
        message: error.message || '此引用对当前操作无效',
        debugInfo: error.debugInfo,
      };
    } else if (errorCode === 'ApiNotFound') {
      return {
        type: 'ApiNotFound',
        code: errorCode,
        message: error.message || '找不到该 API',
        debugInfo: error.debugInfo,
      };
    } else if (errorCode === 'GeneralException') {
      return {
        type: 'GeneralException',
        code: errorCode,
        message: error.message || '处理请求时发生内部错误',
        debugInfo: error.debugInfo,
      };
    }
  }
  
  // 网络错误
  if (error?.name === 'NetworkError' || error?.message?.includes('network')) {
    return {
      type: 'NetworkError',
      message: error.message || '网络错误',
    };
  }
  
  // 其他未知错误
  return {
    type: 'UnknownError',
    message: error instanceof Error ? error.message : String(error),
    stackTrace: error instanceof Error ? error.stack : undefined,
  };
}

/**
 * 保存错误报告
 */
function saveErrorReport(report: ErrorReport): void {
  try {
    // 使用 localStorage 保存错误报告
    const key = `error-report-${report.timestamp}`;
    localStorage.setItem(key, JSON.stringify(report));
    
    // 维护错误报告索引
    const indexKey = 'error-reports-index';
    const indexData = localStorage.getItem(indexKey);
    const index = indexData ? JSON.parse(indexData) : [];
    index.push({
      key,
      timestamp: report.timestamp,
      testCaseId: report.testCaseId,
      errorType: report.errorType,
      hostType: report.hostType,
    });
    
    // 只保留最近 1000 条错误报告
    if (index.length > 1000) {
      const oldestKey = index.shift().key;
      localStorage.removeItem(oldestKey);
    }
    
    localStorage.setItem(indexKey, JSON.stringify(index));
    
    if (DEBUG_EXECUTOR) {
      console.log('[Error Report] Saved:', key);
    }
  } catch (error) {
    console.error('[Error Report] Failed to save:', error);
  }
}

/**
 * 获取 Office 版本信息
 */
function getOfficeVersion(): string {
  try {
    const g = globalThis as any;
    if (g.Office?.context?.diagnostics) {
      return g.Office.context.diagnostics.version || 'unknown';
    }
  } catch {}
  return 'unknown';
}

function normalizeOfficeCode(input: string): string {
  let code = input ?? "";

  // 去掉常见的 fenced code block 包裹
  code = code.replace(/^\s*```[a-zA-Z]*\s*\n/, "");
  code = code.replace(/\n\s*```\s*$/, "");

  // 移除“占位符”行（例如模型输出的 ...）
  code = code.replace(/^\s*\.\.\.\s*$/gm, "");

  // 如果几乎没有真实换行但包含大量 \\n，说明上游把转义序列当成了代码文本（典型：直接从 SSE 原文抽取）
  // 启发式：真实换行少于 3 行，且出现了 \\n
  const realLineCount = code.split("\n").length;
  if (realLineCount < 3 && code.includes("\\n")) {
    code = code.replace(/\\r\\n/g, "\n").replace(/\\n/g, "\n").replace(/\\t/g, "\t").replace(/\\r/g, "\r");
  }

  // 清理多余空行
  code = code.replace(/\n{3,}/g, "\n\n").trim();
  return code;
}

/**
 * 执行 Office.js 代码
 */
export async function executeOfficeCode(
  host: 'word' | 'excel' | 'powerpoint',
  code: string,
  options?: {
    testCaseId?: string;
    userInput?: string;
  }
): Promise<{ success: boolean; result?: any; error?: ExecutionError }> {
  const startTime = Date.now();
  const g = globalThis as any;
  
  if (DEBUG_EXECUTOR) {
    console.group(`[Office Code Executor] Executing ${host} code`);
    console.log('Code length:', code.length);
    console.log('Code preview:', code.substring(0, 150) + '...');
    if (options?.testCaseId) {
      console.log('Test Case ID:', options.testCaseId);
    }
  }
  
  try {
    // 检查全局对象是否可用
    if (typeof g.Office === 'undefined') {
      throw new Error('Office.js 未加载');
    }
    
    const OfficeObj = g.Office;
    const WordObj = typeof g.Word !== 'undefined' ? g.Word : undefined;
    const ExcelObj = typeof g.Excel !== 'undefined' ? g.Excel : undefined;
    const PowerPointObj = typeof g.PowerPoint !== 'undefined' ? g.PowerPoint : undefined;

    // 根据 host 检查对应的全局对象
    const hostGlobals: Record<string, any> = {
      word: WordObj,
      excel: ExcelObj,
      powerpoint: PowerPointObj
    };
    
    if (!hostGlobals[host]) {
      throw new Error(`${host.toUpperCase()} API 未加载`);
    }
    
    // 创建异步函数并执行
    const normalized = normalizeOfficeCode(code);
    
    if (DEBUG_EXECUTOR) {
      console.log('Normalized code:', normalized.substring(0, 200) + '...');
    }
    
    const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;
    const fn = new AsyncFunction('Office', 'Word', 'Excel', 'PowerPoint', normalized);
    
    // 增强错误捕获：包装执行以捕获静默失败
    let executionCompleted = false;
    let lastError: any = null;
    
    const result = await fn(
      OfficeObj,
      WordObj,
      ExcelObj,
      PowerPointObj
    ).then((res: any) => {
      executionCompleted = true;
      return res;
    }).catch((err: any) => {
      lastError = err;
      throw err;
    });
    
    const duration = Date.now() - startTime;
    
    if (DEBUG_EXECUTOR) {
      console.log('✅ Execution successful');
      console.log('Duration:', duration + 'ms');
      console.log('Execution completed flag:', executionCompleted);
      console.log('Result:', result);
      console.log('Result type:', typeof result);
      
      // 警告：如果返回值为undefined且执行时间很短，可能是静默失败
      if (result === undefined && duration < 100) {
        console.warn('⚠️ Warning: Execution succeeded but returned undefined in < 100ms. Possible silent failure.');
      }
      
      console.groupEnd();
    }
    
    return { success: true, result };
  } catch (error) {
    const duration = Date.now() - startTime;
    const executionError = parseOfficeError(error);
    
    if (DEBUG_EXECUTOR) {
      console.error('❌ Execution failed');
      console.error('Duration:', duration + 'ms');
      console.error('Error Type:', executionError.type);
      console.error('Error:', executionError.message);
      if (executionError.debugInfo) {
        console.error('Debug Info:', executionError.debugInfo);
      }
      if (executionError.stackTrace) {
        console.error('Stack Trace:', executionError.stackTrace);
      }
      console.groupEnd();
    } else {
      // 非调试模式下也输出更多信息
      console.error('[Office Code Executor] Error:', executionError.message);
      console.error('[Office Code Executor] Error Type:', executionError.type);
      if (executionError.code) {
        console.error('[Office Code Executor] Error Code:', executionError.code);
      }
    }
    
    // 保存错误报告
    const errorReport: ErrorReport = {
      timestamp: new Date().toISOString(),
      testCaseId: options?.testCaseId,
      hostType: host,
      errorType: executionError.type,
      errorCode: executionError.code,
      errorMessage: executionError.message,
      stackTrace: executionError.stackTrace,
      userInput: options?.userInput,
      generatedCode: normalizeOfficeCode(code),
      context: {
        officeVersion: getOfficeVersion(),
        platform: navigator.platform,
        browserInfo: navigator.userAgent,
      },
    };
    
    saveErrorReport(errorReport);
    
    return { 
      success: false, 
      error: executionError
    };
  }
}

/**
 * 批量执行多个代码块
 */
export async function executeOfficeCodeBlocks(
  blocks: OfficeCodeBlock[],
  options?: {
    testCaseId?: string;
    userInput?: string;
  }
): Promise<Array<{ success: boolean; result?: any; error?: ExecutionError }>> {
  const results: Array<{ success: boolean; result?: any; error?: ExecutionError }> = [];
  
  for (const block of blocks) {
    const result = await executeOfficeCode(block.host, block.code, options);
    results.push(result);
    
    // 如果执行失败，继续执行剩余的代码块
    if (!result.success && DEBUG_EXECUTOR) {
      console.warn(`[Office Code Executor] Failed to execute ${block.host} code, continuing...`);
    }
  }
  
  return results;
}

/**
 * 获取所有错误报告
 */
export function getErrorReports(): ErrorReport[] {
  try {
    const indexKey = 'error-reports-index';
    const indexData = localStorage.getItem(indexKey);
    if (!indexData) return [];
    
    const index = JSON.parse(indexData);
    const reports: ErrorReport[] = [];
    
    for (const item of index) {
      const reportData = localStorage.getItem(item.key);
      if (reportData) {
        reports.push(JSON.parse(reportData));
      }
    }
    
    return reports.sort((a, b) => 
      new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime()
    );
  } catch (error) {
    console.error('[Error Report] Failed to get reports:', error);
    return [];
  }
}

/**
 * 清除所有错误报告
 */
export function clearErrorReports(): void {
  try {
    const indexKey = 'error-reports-index';
    const indexData = localStorage.getItem(indexKey);
    if (!indexData) return;
    
    const index = JSON.parse(indexData);
    for (const item of index) {
      localStorage.removeItem(item.key);
    }
    localStorage.removeItem(indexKey);
    
    console.log('[Error Report] All reports cleared');
  } catch (error) {
    console.error('[Error Report] Failed to clear reports:', error);
  }
}
