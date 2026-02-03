/**
 * 测试执行器
 */

import type { TestCase, TestResult, TestSession, ErrorReport, HostType } from './types';
import { TestLogger } from './TestLogger';
import { executeOfficeCode, extractHiddenOfficeCode } from '@/lib/office/code-executor';

type SSEEvent = { type: string; data: any };

function parseSSEBuffer(buffer: string): { parsed: SSEEvent[]; remaining: string } {
  const parsed: SSEEvent[] = [];
  const chunks = buffer.split("\n\n");
  const remaining = chunks.pop() || "";

  for (const chunk of chunks) {
    if (!chunk.trim()) continue;
    let eventType = "message";
    let dataStr = "";
    for (const line of chunk.split("\n")) {
      if (line.startsWith("event:")) {
        eventType = line.slice(6).trim();
      } else if (line.startsWith("data:")) {
        // 注意：SSE data 可能包含空格，trim 会改变 JSON 字符串；这里只去掉前缀后的一个空格
        const v = line.slice(5);
        dataStr += (v.startsWith(" ") ? v.slice(1) : v);
      }
    }
    if (!dataStr) continue;
    try {
      parsed.push({ type: eventType, data: JSON.parse(dataStr) });
    } catch (e) {
      // 忽略无法解析的事件（避免因单条坏消息中断整套回归）
      console.warn("[TestExecutor] Failed to parse SSE event data:", dataStr.slice(0, 200), e);
    }
  }

  return { parsed, remaining };
}

export class TestExecutor {
  private logger: TestLogger;
  private sessionId: string;
  private hostType: HostType;

  constructor(hostType: HostType) {
    this.hostType = hostType;
    // 测试模式下不使用会话恢复，让 SDK 每次创建新会话
    // 这样避免尝试恢复不存在的会话导致 SDK 进程崩溃
    this.sessionId = `test-${Date.now()}`;
    this.logger = new TestLogger(this.sessionId);
  }

  /**
   * 执行单个测试用例
   */
  async executeTestCase(testCase: TestCase): Promise<TestResult> {
    this.logger.logTestStart(testCase.id, testCase.name);
    
    const result: TestResult = {
      testCaseId: testCase.id,
      status: 'running',
      startTime: new Date().toISOString(),
      logs: [],
    };

    try {
      // 构建完整的 URL，避免相对路径可能的问题
      const baseUrl = window.location.origin;
      const url = `${baseUrl}/api/chat`;

      // 通过 API 发送用户输入到聊天接口
      // 添加简单的重试机制
      let response;
      let lastError;
      
      for (let i = 0; i < 3; i++) {
        try {
          response = await fetch(url, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({
              prompt: testCase.userInput,
              hostType: this.hostType,
              mode: 'Agent', // 明确指定模式
              messages: [],
              // 不传 resume 参数，让 SDK 自动创建新会话
              // 避免尝试恢复不存在的会话导致 SDK 进程崩溃
              testMode: true, // 标记为测试模式
              testCaseId: testCase.id,
              testSessionId: this.sessionId,
            }),
          });
          
          if (response.ok) break;
          // 如果响应不 OK，但在 500-599 之间，可能值得重试
          if (response.status < 500) break;
        } catch (e) {
          lastError = e;
          // 等待一小段时间后重试
          await new Promise(resolve => setTimeout(resolve, 500));
        }
      }

      if (!response) {
        throw lastError || new Error('Network request failed after retries');
      }

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      // 读取 SSE 流并重建最终 assistant 文本，然后提取隐藏 OFFICE-CODE
      const reader = response.body?.getReader();
      const decoder = new TextDecoder();
      let buffer = "";
      const assistantPieces: string[] = [];
      let fullResponsePreview = "";

      if (reader) {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          const chunk = decoder.decode(value, { stream: true });
          buffer += chunk;
          if (fullResponsePreview.length < 4000) {
            fullResponsePreview += chunk;
          }

          const { parsed, remaining } = parseSSEBuffer(buffer);
          buffer = remaining;

          for (const evt of parsed) {
            if (evt.type !== "message") continue;
            const m = evt.data as any;

            // 收集所有可能包含最终文本的来源（stream_event 的 delta、assistant 的 text、result.success 的 result）
            if (m?.type === "stream_event") {
              const delta = m?.event?.delta;
              if (delta?.type === "text_delta" && typeof delta.text === "string") {
                assistantPieces.push(delta.text);
              }
            } else if (m?.type === "assistant") {
              const contentArr = m?.message?.content;
              if (Array.isArray(contentArr)) {
                for (const c of contentArr) {
                  if (c?.type === "text" && typeof c.text === "string") {
                    assistantPieces.push(c.text);
                  }
                }
              }
            } else if (m?.type === "result") {
              if (m?.subtype === "success" && typeof m.result === "string") {
                assistantPieces.push(m.result);
              }
            }
          }
        }
        buffer += decoder.decode();
      }

      // 解析最后残留的 SSE buffer（若有）
      if (buffer.trim()) {
        const { parsed } = parseSSEBuffer(buffer);
        for (const evt of parsed) {
          if (evt.type !== "message") continue;
          const m = evt.data as any;
          if (m?.type === "assistant") {
            const contentArr = m?.message?.content;
            if (Array.isArray(contentArr)) {
              for (const c of contentArr) {
                if (c?.type === "text" && typeof c.text === "string") {
                  assistantPieces.push(c.text);
                }
              }
            }
          } else if (m?.type === "result") {
            if (m?.subtype === "success" && typeof m.result === "string") {
              assistantPieces.push(m.result);
            }
          }
        }
      }

      const assistantText = assistantPieces.join("");
      const blocks = extractHiddenOfficeCode(assistantText);
      const hostBlocks = blocks.filter(b => b.host === this.hostType);
      const uniqueCodes = Array.from(new Set(hostBlocks.map(b => b.code).filter(Boolean)));
      const generatedCode = uniqueCodes.join("\n\n");

      // 前端打印：输入、生成代码、执行结果（便于诊断）
      console.groupCollapsed(`[TestExecutor] ${this.hostType.toUpperCase()} ${testCase.id}`);
      console.log("userInput:", testCase.userInput);
      console.log("assistantTextPreview:", assistantText.slice(0, 400));
      console.log("generatedCodeLength:", generatedCode.length);
      console.log("generatedCodePreview:", generatedCode.slice(0, 400));
      console.groupEnd();

      result.actualCode = generatedCode.trim();
      result.endTime = new Date().toISOString();
      result.duration = new Date(result.endTime).getTime() - new Date(result.startTime).getTime();

      // 简单判断: 如果生成了代码就认为通过
      if (generatedCode.length > 0) {
        result.status = 'passed';
        // 尝试在 Office 环境中执行代码
        const execResult = await executeOfficeCode(this.hostType, generatedCode, {
          testCaseId: testCase.id,
          userInput: testCase.userInput
        });

        // 后端打印：将执行结果回传到 /api/test-report（不影响测试流程）
        try {
          await fetch(`${baseUrl}/api/test-report`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              timestamp: new Date().toISOString(),
              sessionId: this.sessionId,
              testCaseId: testCase.id,
              hostType: this.hostType,
              userInput: testCase.userInput,
              generatedCode,
              execution: execResult,
            }),
          });
        } catch (e) {
          this.logger.warn(
            `回传 test-report 失败（不影响测试结果）: ${e instanceof Error ? e.message : String(e)}`
          );
        }
        
        if (!execResult.success) {
          result.status = 'failed';
          result.error = {
            type: execResult.error?.type || 'UnknownError',
            message: execResult.error?.message || '代码执行失败',
            stackTrace: execResult.error?.stackTrace
          };
        }
      } else {
        result.status = 'failed';
        result.error = {
          type: 'UnknownError',
          message: '未生成 Office 代码',
          debugInfo: { preview: fullResponsePreview, assistantTextPreview: assistantText.slice(0, 1000) } // 记录预览以便调试
        };
        this.logger.warn(`未生成代码。SSE 预览: ${fullResponsePreview.substring(0, 200)}...`);
      }

    } catch (error) {
      result.status = 'failed';
      result.endTime = new Date().toISOString();
      result.duration = new Date(result.endTime).getTime() - new Date(result.startTime).getTime();
      
      result.error = {
        type: 'NetworkError',
        message: error instanceof Error ? error.message : String(error),
        stackTrace: error instanceof Error ? error.stack : undefined,
      };

      // 保存错误报告
      const errorReport: ErrorReport = {
        timestamp: new Date().toISOString(),
        testCaseId: testCase.id,
        hostType: this.hostType,
        errorType: result.error.type,
        errorMessage: result.error.message,
        stackTrace: result.error.stackTrace,
        userInput: testCase.userInput,
        generatedCode: result.actualCode || '',
        context: {
          officeVersion: 'unknown',
          platform: navigator.platform,
          browserInfo: navigator.userAgent,
        },
      };

      await this.logger.saveErrorReport(errorReport);
    }

    this.logger.logTestComplete(result);
    return result;
  }

  /**
   * 执行测试套件
   */
  async executeTestSuite(
    testCases: TestCase[],
    onProgress?: (current: number, total: number) => void
  ): Promise<TestSession> {
    const session: TestSession = {
      id: this.sessionId,
      hostType: this.hostType,
      startTime: new Date().toISOString(),
      results: [],
      summary: {
        total: testCases.length,
        passed: 0,
        failed: 0,
        skipped: 0,
        errorRate: 0,
      },
    };

    this.logger.info(`开始执行测试套件: ${testCases.length} 个测试用例`);

    for (let i = 0; i < testCases.length; i++) {
      const testCase = testCases[i];
      
      if (onProgress) {
        onProgress(i + 1, testCases.length);
      }

      const result = await this.executeTestCase(testCase);
      session.results.push(result);

      // 更新统计
      if (result.status === 'passed') {
        session.summary.passed++;
      } else if (result.status === 'failed') {
        session.summary.failed++;
      } else if (result.status === 'skipped') {
        session.summary.skipped++;
      }

      // 添加延迟避免请求过快
      await new Promise(resolve => setTimeout(resolve, 1000));
    }

    session.endTime = new Date().toISOString();
    session.summary.errorRate = session.summary.failed / session.summary.total;

    this.logger.logSessionSummary(session);

    return session;
  }

  /**
   * 获取日志
   */
  getLogger(): TestLogger {
    return this.logger;
  }
}
