import { NextRequest } from "next/server";
import { onNewTask } from "@/lib/office/mcp-server";

export const runtime = 'nodejs';
export const dynamic = 'force-dynamic';

/**
 * SSE (Server-Sent Events) 推送任务到前端
 * 
 * 替代轮询机制，实现即时推送：
 * - 零延迟：任务产生时立即推送
 * - 低开销：单个长连接，无需每 100ms 轮询
 * - 自动重连：连接断开后前端自动重连
 */
export async function GET(request: NextRequest) {
  const encoder = new TextEncoder();
  
  console.log('[SSE] New client connection established');
  
  // 创建 SSE 流
  const stream = new ReadableStream({
    start(controller) {
      // 发送初始连接消息
      try {
        const connectMessage = `data: ${JSON.stringify({ type: 'connected', timestamp: Date.now() })}\n\n`;
        controller.enqueue(encoder.encode(connectMessage));
        console.log('[SSE] Sent connection confirmation');
      } catch (error) {
        console.error('[SSE] Error sending connection message:', error);
      }
      
      // 监听新任务事件
      const cleanup = onNewTask((task) => {
        try {
          const message = `data: ${JSON.stringify({
            type: 'task',
            correlationId: task.correlationId,
            host: task.host,
            code: task.code,
            description: task.description,
            timestamp: Date.now()
          })}\n\n`;
          
          controller.enqueue(encoder.encode(message));
          console.log(`[SSE] Pushed task ${task.correlationId} (${task.host}) to client`);
        } catch (error) {
          console.error('[SSE] Error pushing task:', error);
        }
      });
      
      // 定期发送心跳（每 30 秒）
      // 保持连接活跃，防止代理/防火墙超时断开
      const heartbeatInterval = setInterval(() => {
        try {
          controller.enqueue(encoder.encode(`: heartbeat ${Date.now()}\n\n`));
        } catch (error) {
          console.error('[SSE] Heartbeat error:', error);
          clearInterval(heartbeatInterval);
        }
      }, 30000);
      
      // 清理函数：连接关闭时执行
      request.signal.addEventListener('abort', () => {
        clearInterval(heartbeatInterval);
        cleanup();
        try {
          controller.close();
        } catch (error) {
          // 连接可能已关闭，忽略错误
        }
        console.log('[SSE] Connection closed, cleaned up resources');
      });
    }
  });
  
  // 返回 SSE 响应
  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      'Connection': 'keep-alive',
      'X-Accel-Buffering': 'no', // 禁用 nginx 缓冲
    }
  });
}
