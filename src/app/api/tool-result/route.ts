import { NextRequest, NextResponse } from "next/server";
import { 
  submitExecutionResult, 
  getAllPendingExecutions,
  ExecutionResult 
} from "@/lib/office/mcp-server";

// 存储待处理的工具结果（用于 office_action 事件的旧机制）
const pendingResults = new Map<string, { result: unknown; timestamp: number }>();

// 清理过期结果（超过 5 分钟）
function cleanupOldResults() {
  const now = Date.now();
  const fiveMinutes = 5 * 60 * 1000;
  
  for (const [key, value] of pendingResults) {
    if (now - value.timestamp > fiveMinutes) {
      pendingResults.delete(key);
    }
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { correlationId, result, source } = body;
    
    if (!correlationId) {
      return NextResponse.json(
        { error: "correlationId is required" },
        { status: 400 }
      );
    }
    
    // 判断来源：MCP Tool 还是旧的 office_action 机制
    if (source === 'mcp_tool') {
      // 新机制：MCP Tool 执行结果
      const executionResult: ExecutionResult = {
        success: result.success,
        data: result.data,
        error: result.error,
        timestamp: Date.now()
      };
      
      const submitted = submitExecutionResult(correlationId, executionResult);
      
      if (!submitted) {
        console.warn(`[API/tool-result] Unknown MCP correlationId: ${correlationId}`);
      } else {
        console.log(`[API/tool-result] MCP result submitted for ${correlationId}, success: ${result.success}`);
      }
      
      return NextResponse.json({ success: true });
    } else {
      // 旧机制：office_action 事件结果
      pendingResults.set(correlationId, {
        result,
        timestamp: Date.now(),
      });
      
      // 清理过期结果
      cleanupOldResults();
      
      console.log(`[API/tool-result] Received result for ${correlationId}`);
      
      return NextResponse.json({ success: true });
    }
  } catch (error) {
    console.error("[API/tool-result] Error:", error);
    return NextResponse.json(
      { error: (error as Error).message },
      { status: 500 }
    );
  }
}

export async function GET(request: NextRequest) {
  const correlationId = request.nextUrl.searchParams.get("correlationId");
  const action = request.nextUrl.searchParams.get("action");
  
  // 新功能：获取所有待执行的 MCP 任务
  if (action === "pending_executions") {
    const pending = getAllPendingExecutions();
    return NextResponse.json({ 
      executions: pending.map(({ correlationId, execution }) => ({
        correlationId,
        host: execution.host,
        code: execution.code,
        description: execution.description
      }))
    });
  }
  
  // 旧机制：根据 correlationId 获取结果
  if (!correlationId) {
    return NextResponse.json(
      { error: "correlationId is required" },
      { status: 400 }
    );
  }
  
  const pending = pendingResults.get(correlationId);
  
  if (pending) {
    pendingResults.delete(correlationId);
    return NextResponse.json({ 
      found: true, 
      result: pending.result 
    });
  }
  
  return NextResponse.json({ found: false });
}
