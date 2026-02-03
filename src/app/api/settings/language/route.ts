import { NextRequest, NextResponse } from "next/server";
import { promises as fs } from "fs";
import path from "path";

export async function POST(request: NextRequest) {
  try {
    const { language } = await request.json();

    if (!language || typeof language !== "string") {
      return NextResponse.json(
        { error: "Invalid language parameter" },
        { status: 400 }
      );
    }

    // 构建 .claude/settings.local.json 的路径
    const settingsPath = path.join(process.cwd(), ".claude", "settings.local.json");

    // 读取现有设置
    let settings: any = {};
    try {
      const fileContent = await fs.readFile(settingsPath, "utf-8");
      settings = JSON.parse(fileContent);
    } catch (error) {
      // 如果文件不存在，创建新的设置对象
      console.log("[API/settings/language] Creating new settings file");
      settings = {};
    }

    // 更新语言设置（保留其他设置）
    settings.language = language;

    // 确保 .claude 目录存在
    const claudeDir = path.join(process.cwd(), ".claude");
    try {
      await fs.access(claudeDir);
    } catch {
      await fs.mkdir(claudeDir, { recursive: true });
    }

    // 写入文件
    await fs.writeFile(settingsPath, JSON.stringify(settings, null, 2), "utf-8");

    console.log("[API/settings/language] Language setting saved:", language);

    return NextResponse.json({ success: true, language });
  } catch (error) {
    console.error("[API/settings/language] Error:", error);
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "Unknown error" },
      { status: 500 }
    );
  }
}

export async function GET() {
  try {
    // 构建 .claude/settings.local.json 的路径
    const settingsPath = path.join(process.cwd(), ".claude", "settings.local.json");

    // 读取设置
    const fileContent = await fs.readFile(settingsPath, "utf-8");
    const settings = JSON.parse(fileContent);

    return NextResponse.json({ language: settings.language || "default" });
  } catch (error) {
    console.error("[API/settings/language] Error reading settings:", error);
    return NextResponse.json({ language: "default" });
  }
}
