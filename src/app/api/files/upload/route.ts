import { NextRequest, NextResponse } from "next/server";
import { writeFile, mkdir } from "fs/promises";
import { existsSync } from "fs";
import path from "path";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const sessionId = formData.get("sessionId") as string | null;

    if (!file) {
      return NextResponse.json(
        { error: "没有文件" },
        { status: 400 }
      );
    }

    // 验证文件类型
    const allowedExtensions = [
      ".txt", ".pdf", ".doc", ".docx", 
      ".xls", ".xlsx", ".csv", ".json",
      ".png", ".jpg", ".jpeg"
    ];
    const fileExtension = path.extname(file.name).toLowerCase();
    if (!allowedExtensions.includes(fileExtension)) {
      return NextResponse.json(
        { error: `不支持的文件类型: ${fileExtension}` },
        { status: 400 }
      );
    }

    // 验证文件大小（最大 10MB）
    const maxSize = 10 * 1024 * 1024;
    if (file.size > maxSize) {
      return NextResponse.json(
        { error: "文件大小超过 10MB" },
        { status: 400 }
      );
    }

    // 获取项目根目录
    const projectRoot = process.cwd();
    
    // 确定保存路径
    let uploadDir: string;
    if (sessionId) {
      uploadDir = path.join(projectRoot, "workspace", "sessions", sessionId, "uploads");
    } else {
      uploadDir = path.join(projectRoot, "workspace", "temp_uploads");
    }

    // 创建目录（如果不存在）
    if (!existsSync(uploadDir)) {
      await mkdir(uploadDir, { recursive: true });
    }

    // 生成唯一文件名
    const timestamp = Date.now();
    const sanitizedFileName = file.name.replace(/[^a-zA-Z0-9._-]/g, "_");
    const fileName = `${timestamp}_${sanitizedFileName}`;
    const filePath = path.join(uploadDir, fileName);

    // 将文件写入磁盘
    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    await writeFile(filePath, buffer);

    console.log("[API/files/upload] File saved:", {
      sessionId: sessionId || "temp",
      fileName,
      size: file.size,
      path: filePath
    });

    return NextResponse.json({
      success: true,
      fileId: fileName,
      fileName: file.name,
      size: file.size,
      path: sessionId 
        ? `workspace/sessions/${sessionId}/uploads/${fileName}`
        : `workspace/temp_uploads/${fileName}`,
    });
  } catch (error) {
    console.error("[API/files/upload] Error:", error);
    return NextResponse.json(
      { error: `上传失败: ${(error as Error).message}` },
      { status: 500 }
    );
  }
}
