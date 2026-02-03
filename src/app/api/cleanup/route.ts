import { NextRequest, NextResponse } from "next/server";
import { runCleanup } from "@/lib/cleanup";

/**
 * 手动触发清理任务的 API 端点
 * 可以通过 cron 任务或手动调用
 */
export async function POST(request: NextRequest) {
  try {
    // 验证请求来源（可选：添加 API Key 验证）
    const authHeader = request.headers.get("authorization");
    const apiKey = process.env.CLEANUP_API_KEY;

    // 如果设置了 API Key，则验证
    if (apiKey && authHeader !== `Bearer ${apiKey}`) {
      return NextResponse.json(
        { error: "未授权" },
        { status: 401 }
      );
    }

    console.log("[API/cleanup] Starting cleanup task...");
    await runCleanup();

    return NextResponse.json({
      success: true,
      message: "清理任务已完成"
    });
  } catch (error) {
    console.error("[API/cleanup] Error:", error);
    return NextResponse.json(
      { error: `清理任务失败: ${(error as Error).message}` },
      { status: 500 }
    );
  }
}

/**
 * GET 请求返回清理任务状态
 */
export async function GET(request: NextRequest) {
  return NextResponse.json({
    message: "清理任务 API 端点",
    usage: "POST 请求触发清理任务",
    auth: process.env.CLEANUP_API_KEY ? "需要 Bearer Token" : "无需认证"
  });
}
