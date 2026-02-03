"use client";

export type OfficeHostType = "excel" | "word" | "powerpoint" | "unknown";

/**
 * 检测当前 Office 宿主类型
 * 在 Office 加载项初始化后调用
 */
export function getOfficeHost(): OfficeHostType {
  if (typeof Office === "undefined" || !Office.context) {
    return "unknown";
  }
  
  switch (Office.context.host) {
    case Office.HostType.Excel:
      return "excel";
    case Office.HostType.Word:
      return "word";
    case Office.HostType.PowerPoint:
      return "powerpoint";
    default:
      return "unknown";
  }
}

export function isExcelHost(): boolean {
  return getOfficeHost() === "excel";
}

export function isWordHost(): boolean {
  return getOfficeHost() === "word";
}

export function isPowerPointHost(): boolean {
  return getOfficeHost() === "powerpoint";
}

export function isOfficeHost(): boolean {
  return getOfficeHost() !== "unknown";
}

export function getOfficeHostDisplayName(): string {
  const host = getOfficeHost();
  switch (host) {
    case "excel": return "Excel";
    case "word": return "Word";
    case "powerpoint": return "PowerPoint";
    default: return "未知应用";
  }
}
