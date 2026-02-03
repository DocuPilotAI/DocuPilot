import { NextRequest, NextResponse } from "next/server";
import { unlink } from "fs/promises";
import { existsSync } from "fs";
import path from "path";

export async function DELETE(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const sessionId = searchParams.get("sessionId");
    const fileId = searchParams.get("fileId");

    if (!sessionId || !fileId) {
      return NextResponse.json(
        { error: "缺少 sessionId 或 fileId 参数" },
        { status: 400 }
      );
    }

    const projectRoot = process.cwd();
    const filePath = path.join(
      projectRoot,
      "workspace",
      "sessions",
      sessionId,
      "uploads",
      fileId
    );

    // 验证文件路径安全性（防止路径遍历攻击）
    const normalizedPath = path.normalize(filePath);
    const workspaceDir = path.join(projectRoot, "workspace", "sessions", sessionId);
    if (!normalizedPath.startsWith(workspaceDir)) {
      return NextResponse.json(
        { error: "非法的文件路径" },
        { status: 403 }
      );
    }

    if (!existsSync(filePath)) {
      return NextResponse.json(
        { error: "文件不存在" },
        { status: 404 }
      );
    }

    await unlink(filePath);

    console.log("[API/files/delete] File deleted:", {
      sessionId,
      fileId,
      path: filePath
    });

    return NextResponse.json({
      success: true,
      message: "文件已删除"
    });
  } catch (error) {
    console.error("[API/files/delete] Error:", error);
    return NextResponse.json(
      { error: `删除文件失败: ${(error as Error).message}` },
      { status: 500 }
    );
  }
}
