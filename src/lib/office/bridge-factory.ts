"use client";

import { getOfficeHost, OfficeHostType } from "./host-detector";

const DEBUG_OFFICE =
  process.env.NODE_ENV === "development" ||
  (typeof process !== "undefined" && process.env?.NEXT_PUBLIC_DEBUG_CHAT === "1");

function payloadSummary(payload: Record<string, unknown>): string {
  if (payload.text != null) {
    const s = String(payload.text);
    return `textLen=${s.length}`;
  }
  return `keys=${Object.keys(payload).join(",")}`;
}

export interface OfficeBridge {
  getHostType(): OfficeHostType;
  isHostAvailable(): boolean;
  getSelectedContent(): Promise<unknown>;
  writeContent(content: unknown): Promise<{ success: boolean; error?: string }>;
  handleAction(action: string, payload: Record<string, unknown>): Promise<unknown>;
}

// Excel Bridge Implementation
class ExcelBridge implements OfficeBridge {
  getHostType(): OfficeHostType {
    return "excel";
  }

  isHostAvailable(): boolean {
    return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Excel;
  }

  async getSelectedContent(): Promise<unknown> {
    if (!this.isHostAvailable()) {
      return { error: "Excel not available" };
    }
    
    return new Promise((resolve) => {
      // @ts-expect-error Excel API
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address", "formulas"]);
        await context.sync();
        
        resolve({
          address: range.address,
          values: range.values,
          formulas: range.formulas,
        });
      }).catch((error: Error) => {
        resolve({ error: error.message });
      });
    });
  }

  async writeContent(content: unknown): Promise<{ success: boolean; error?: string }> {
    if (!this.isHostAvailable()) {
      return { success: false, error: "Excel not available" };
    }

    return new Promise((resolve) => {
      // @ts-expect-error Excel API
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        if (Array.isArray(content)) {
          range.values = content;
        }
        await context.sync();
        resolve({ success: true });
      }).catch((error: Error) => {
        resolve({ success: false, error: error.message });
      });
    });
  }

  async handleAction(action: string, payload: Record<string, unknown>): Promise<unknown> {
    switch (action) {
      case "read_range":
        return this.getSelectedContent();
      case "write_range":
        return this.writeContent(payload.values);
      case "get_selection":
        return this.getSelectedContent();
      default:
        return { error: `Unknown action: ${action}` };
    }
  }
}

// Word Bridge Implementation
class WordBridge implements OfficeBridge {
  getHostType(): OfficeHostType {
    return "word";
  }

  isHostAvailable(): boolean {
    return typeof Office !== "undefined" && Office.context?.host === Office.HostType.Word;
  }

  async getSelectedContent(): Promise<unknown> {
    if (!this.isHostAvailable()) {
      return { error: "Word not available" };
    }
    
    return new Promise((resolve) => {
      // @ts-expect-error Word API
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
        
        resolve({
          text: selection.text,
        });
      }).catch((error: Error) => {
        resolve({ error: error.message });
      });
    });
  }

  async writeContent(content: unknown): Promise<{ success: boolean; error?: string }> {
    if (DEBUG_OFFICE) {
      const s = String(content);
      console.log("[WordBridge] writeContent", { contentLen: s.length });
    }
    if (!this.isHostAvailable()) {
      const out = { success: false as const, error: "Word not available" };
      if (DEBUG_OFFICE) console.log("[WordBridge] writeContent result", out);
      return out;
    }

    return new Promise((resolve) => {
      // @ts-expect-error Word API
      Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.insertText(String(content), "Replace");
        await context.sync();
        const out = { success: true as const };
        if (DEBUG_OFFICE) console.log("[WordBridge] writeContent result", out);
        resolve(out);
      }).catch((error: Error) => {
        const out = { success: false as const, error: error.message };
        if (DEBUG_OFFICE) console.log("[WordBridge] writeContent result", out);
        resolve(out);
      });
    });
  }

  async handleAction(action: string, payload: Record<string, unknown>): Promise<unknown> {
    if (DEBUG_OFFICE) {
      console.log("[WordBridge] handleAction", { action, payloadSummary: payloadSummary(payload) });
    }
    switch (action) {
      case "read_document":
        return this.getSelectedContent();
      case "insert_text":
        return this.writeContent(payload.text);
      case "insert_paragraph":
        return this.insertParagraph(payload.text as string);
      default:
        return { error: `Unknown action: ${action}` };
    }
  }

  private async insertParagraph(text: string): Promise<{ success: boolean; error?: string }> {
    if (DEBUG_OFFICE) {
      console.log("[WordBridge] insertParagraph", { textLen: text.length });
    }
    if (!this.isHostAvailable()) {
      const out = { success: false as const, error: "Word not available" };
      if (DEBUG_OFFICE) console.log("[WordBridge] insertParagraph result", out);
      return out;
    }

    return new Promise((resolve) => {
      // @ts-expect-error Word API
      Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(text, "End");
        paragraph.load("text");
        await context.sync();
        const out = { success: true as const };
        if (DEBUG_OFFICE) console.log("[WordBridge] insertParagraph result", out);
        resolve(out);
      }).catch((error: Error) => {
        const out = { success: false as const, error: error.message };
        if (DEBUG_OFFICE) console.log("[WordBridge] insertParagraph result", out);
        resolve(out);
      });
    });
  }
}

// PowerPoint Bridge Implementation
class PowerPointBridge implements OfficeBridge {
  getHostType(): OfficeHostType {
    return "powerpoint";
  }

  isHostAvailable(): boolean {
    return typeof Office !== "undefined" && Office.context?.host === Office.HostType.PowerPoint;
  }

  async getSelectedContent(): Promise<unknown> {
    if (!this.isHostAvailable()) {
      return { error: "PowerPoint not available" };
    }
    
    return new Promise((resolve) => {
      // @ts-expect-error PowerPoint API
      PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        resolve({
          slideCount: slides.items.length,
        });
      }).catch((error: Error) => {
        resolve({ error: error.message });
      });
    });
  }

  async writeContent(content: unknown): Promise<{ success: boolean; error?: string }> {
    return { success: false, error: "Use specific PowerPoint actions instead" };
  }

  async handleAction(action: string, payload: Record<string, unknown>): Promise<unknown> {
    switch (action) {
      case "get_slides":
        return this.getSelectedContent();
      case "add_slide":
        return this.addSlide();
      default:
        return { error: `Unknown action: ${action}` };
    }
  }

  private async addSlide(): Promise<{ success: boolean; error?: string }> {
    if (!this.isHostAvailable()) {
      return { success: false, error: "PowerPoint not available" };
    }

    return new Promise((resolve) => {
      // @ts-expect-error PowerPoint API
      PowerPoint.run(async (context) => {
        context.presentation.slides.add();
        await context.sync();
        resolve({ success: true });
      }).catch((error: Error) => {
        resolve({ success: false, error: error.message });
      });
    });
  }
}

// Null Bridge for unknown hosts
class NullBridge implements OfficeBridge {
  getHostType(): OfficeHostType {
    return "unknown";
  }

  isHostAvailable(): boolean {
    return false;
  }

  async getSelectedContent(): Promise<unknown> {
    return { error: "No Office host available" };
  }

  async writeContent(): Promise<{ success: boolean; error?: string }> {
    return { success: false, error: "No Office host available" };
  }

  async handleAction(): Promise<unknown> {
    return { error: "No Office host available" };
  }
}

/**
 * 根据当前宿主自动获取对应的桥接实例
 */
export function getBridge(): OfficeBridge {
  const hostType = getOfficeHost();
  return getBridgeByType(hostType);
}

/**
 * 根据指定类型获取桥接实例
 */
export function getBridgeByType(hostType: OfficeHostType): OfficeBridge {
  switch (hostType) {
    case "excel":
      return new ExcelBridge();
    case "word":
      return new WordBridge();
    case "powerpoint":
      return new PowerPointBridge();
    default:
      return new NullBridge();
  }
}

/**
 * 统一的 Office 操作入口
 * 根据当前宿主自动路由到对应的处理函数
 */
export async function handleOfficeAction(
  action: string,
  payload: Record<string, unknown>,
  hostType?: OfficeHostType
): Promise<unknown> {
  const type = hostType ?? getOfficeHost();
  if (DEBUG_OFFICE) {
    console.log("[handleOfficeAction]", { action, hostType: type, payloadSummary: payloadSummary(payload) });
  }
  const bridge = getBridgeByType(type);
  const result = await bridge.handleAction(action, payload);
  if (DEBUG_OFFICE && result && typeof result === "object" && "success" in result) {
    const r = result as { success?: boolean; error?: string };
    console.log("[handleOfficeAction] result", { action, success: r.success, error: r.error ?? "—" });
  }
  return result;
}
