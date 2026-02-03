import { NextRequest, NextResponse } from "next/server";

/**
 * 前端测试执行结果回传接口（仅用于日志/诊断）
 * 目的：让后端也能打印每条用例的输入、生成代码、执行结果。
 */
export async function POST(request: NextRequest) {
  try {
    const body = await request.json();

    const {
      sessionId,
      testCaseId,
      hostType,
      userInput,
      generatedCode,
      execution,
      timestamp,
    } = body ?? {};

    console.log("[API/test-report]", {
      timestamp: timestamp || new Date().toISOString(),
      sessionId,
      testCaseId,
      hostType,
      userInput,
      generatedCodePreview:
        typeof generatedCode === "string" ? generatedCode.slice(0, 300) : undefined,
      generatedCodeLength: typeof generatedCode === "string" ? generatedCode.length : undefined,
      execution,
    });

    return NextResponse.json({ success: true });
  } catch (error) {
    console.error("[API/test-report] Error:", error);
    return NextResponse.json({ error: (error as Error).message }, { status: 500 });
  }
}

