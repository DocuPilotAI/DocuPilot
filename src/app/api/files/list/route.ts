import { NextRequest, NextResponse } from "next/server";
import { readdir, stat } from "fs/promises";
import { existsSync } from "fs";
import path from "path";

export async function GET(request: NextRequest) {
  try {
    const { searchParams } = new URL(request.url);
    const sessionId = searchParams.get("sessionId");

    if (!sessionId) {
      return NextResponse.json(
        { error: "缺少 sessionId 参数" },
        { status: 400 }
      );
    }

    const projectRoot = process.cwd();
    const uploadDir = path.join(projectRoot, "workspace", "sessions", sessionId, "uploads");

    if (!existsSync(uploadDir)) {
      return NextResponse.json({ files: [] });
    }

    const fileNames = await readdir(uploadDir);
    const files = await Promise.all(
      fileNames.map(async (fileName) => {
        const filePath = path.join(uploadDir, fileName);
        const stats = await stat(filePath);
        
        // 提取原始文件名（去掉时间戳前缀）
        const originalName = fileName.replace(/^\d+_/, "");
        
        return {
          fileId: fileName,
          fileName: originalName,
          size: stats.size,
          uploadTime: stats.mtime.toISOString(),
          path: `workspace/sessions/${sessionId}/uploads/${fileName}`,
        };
      })
    );

    return NextResponse.json({ files });
  } catch (error) {
    console.error("[API/files/list] Error:", error);
    return NextResponse.json(
      { error: `获取文件列表失败: ${(error as Error).message}` },
      { status: 500 }
    );
  }
}
