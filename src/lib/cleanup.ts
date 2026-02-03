import { readdir, stat, unlink, rmdir } from "fs/promises";
import { existsSync } from "fs";
import path from "path";

/**
 * 清理临时上传文件（超过指定时间的文件）
 * @param maxAgeHours 文件最大保留时间（小时）
 */
export async function cleanupTempUploads(maxAgeHours: number = 1): Promise<void> {
  const projectRoot = process.cwd();
  const tempDir = path.join(projectRoot, "workspace", "temp_uploads");

  if (!existsSync(tempDir)) {
    return;
  }

  const now = Date.now();
  const maxAgeMs = maxAgeHours * 60 * 60 * 1000;

  try {
    const files = await readdir(tempDir);
    let deletedCount = 0;

    for (const file of files) {
      const filePath = path.join(tempDir, file);
      const stats = await stat(filePath);

      // 检查文件修改时间
      const fileAge = now - stats.mtimeMs;
      if (fileAge > maxAgeMs) {
        await unlink(filePath);
        deletedCount++;
        console.log(`[Cleanup] Deleted temp file: ${file} (age: ${Math.round(fileAge / 3600000)}h)`);
      }
    }

    console.log(`[Cleanup] Temp files cleanup completed. Deleted ${deletedCount} files.`);
  } catch (error) {
    console.error("[Cleanup] Error cleaning temp uploads:", error);
  }
}

/**
 * 清理过期的会话目录（超过指定时间未修改的会话）
 * @param maxAgeHours 会话最大保留时间（小时）
 */
export async function cleanupExpiredSessions(maxAgeHours: number = 24): Promise<void> {
  const projectRoot = process.cwd();
  const sessionsDir = path.join(projectRoot, "workspace", "sessions");

  if (!existsSync(sessionsDir)) {
    return;
  }

  const now = Date.now();
  const maxAgeMs = maxAgeHours * 60 * 60 * 1000;

  try {
    const sessionIds = await readdir(sessionsDir);
    let deletedCount = 0;

    for (const sessionId of sessionIds) {
      const sessionPath = path.join(sessionsDir, sessionId);
      const stats = await stat(sessionPath);

      // 检查目录修改时间
      const sessionAge = now - stats.mtimeMs;
      if (sessionAge > maxAgeMs) {
        await deleteDirectoryRecursive(sessionPath);
        deletedCount++;
        console.log(`[Cleanup] Deleted expired session: ${sessionId} (age: ${Math.round(sessionAge / 3600000)}h)`);
      }
    }

    console.log(`[Cleanup] Sessions cleanup completed. Deleted ${deletedCount} sessions.`);
  } catch (error) {
    console.error("[Cleanup] Error cleaning expired sessions:", error);
  }
}

/**
 * 递归删除目录及其内容
 */
async function deleteDirectoryRecursive(dirPath: string): Promise<void> {
  if (!existsSync(dirPath)) {
    return;
  }

  const files = await readdir(dirPath);

  for (const file of files) {
    const filePath = path.join(dirPath, file);
    const stats = await stat(filePath);

    if (stats.isDirectory()) {
      await deleteDirectoryRecursive(filePath);
    } else {
      await unlink(filePath);
    }
  }

  await rmdir(dirPath);
}

/**
 * 执行完整的清理任务
 */
export async function runCleanup(): Promise<void> {
  console.log("[Cleanup] Starting cleanup task...");
  
  await cleanupTempUploads(1); // 清理超过 1 小时的临时文件
  await cleanupExpiredSessions(24); // 清理超过 24 小时的会话
  
  console.log("[Cleanup] Cleanup task completed.");
}

// 如果直接运行此脚本，执行清理任务
if (require.main === module) {
  runCleanup().catch(console.error);
}
