#!/usr/bin/env node

/**
 * SSE 推送功能测试脚本
 * 
 * 测试内容：
 * 1. SSE 连接建立
 * 2. 任务推送
 * 3. 降级到轮询
 */

console.log("=".repeat(60));
console.log("SSE 任务推送测试");
console.log("=".repeat(60));
console.log();

// 测试配置
const SERVER_URL = 'https://localhost:3000';
const SSE_ENDPOINT = `${SERVER_URL}/api/task-stream`;

// Node.js HTTPS 自签名证书支持
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

/**
 * 测试 SSE 连接
 */
async function testSSEConnection() {
  console.log("📡 测试 SSE 连接...");
  
  return new Promise((resolve, reject) => {
    try {
      // Node.js 中使用 fetch 模拟 EventSource
      const controller = new AbortController();
      const timeoutId = setTimeout(() => {
        controller.abort();
        reject(new Error('连接超时'));
      }, 10000);
      
      fetch(SSE_ENDPOINT, {
        signal: controller.signal,
        headers: {
          'Accept': 'text/event-stream'
        }
      }).then(async response => {
        clearTimeout(timeoutId);
        
        if (!response.ok) {
          reject(new Error(`HTTP ${response.status}: ${response.statusText}`));
          return;
        }
        
        if (!response.body) {
          reject(new Error('响应体为空'));
          return;
        }
        
        console.log("✅ SSE 连接成功");
        console.log(`   - Content-Type: ${response.headers.get('Content-Type')}`);
        console.log(`   - Cache-Control: ${response.headers.get('Cache-Control')}`);
        
        // 读取前几个消息
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let receivedMessages = 0;
        let hasConnectedMessage = false;
        
        const readTimeout = setTimeout(() => {
          controller.abort();
        }, 5000);
        
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            
            for (const line of lines) {
              if (line.startsWith('data:')) {
                const data = line.substring(5).trim();
                try {
                  const json = JSON.parse(data);
                  receivedMessages++;
                  
                  if (json.type === 'connected') {
                    hasConnectedMessage = true;
                    console.log("✅ 收到连接确认消息");
                  } else if (json.type === 'task') {
                    console.log(`📦 收到任务: ${json.correlationId} (${json.host})`);
                  }
                } catch (e) {
                  // 可能是其他格式的消息
                }
              } else if (line.startsWith(':')) {
                console.log(`💓 收到心跳`);
              }
            }
            
            // 收到连接消息后可以结束测试
            if (hasConnectedMessage) {
              clearTimeout(readTimeout);
              controller.abort();
              break;
            }
          }
        } catch (error) {
          if (error.name !== 'AbortError') {
            console.error('读取错误:', error);
          }
        }
        
        console.log(`\n📊 测试结果:`);
        console.log(`   - 收到消息数: ${receivedMessages}`);
        console.log(`   - 连接消息: ${hasConnectedMessage ? '✅' : '❌'}`);
        
        resolve({
          success: hasConnectedMessage,
          messagesReceived: receivedMessages
        });
      }).catch(error => {
        clearTimeout(timeoutId);
        reject(error);
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * 测试轮询端点（降级方案）
 */
async function testPollingEndpoint() {
  console.log("\n📊 测试轮询端点（降级方案）...");
  
  try {
    const response = await fetch(`${SERVER_URL}/api/tool-result?action=pending_executions`);
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    const data = await response.json();
    
    console.log("✅ 轮询端点正常");
    console.log(`   - 待执行任务数: ${data.executions?.length || 0}`);
    
    return { success: true };
  } catch (error) {
    console.error("❌ 轮询端点失败:", error.message);
    return { success: false, error };
  }
}

/**
 * 主测试函数
 */
async function main() {
  let allTestsPassed = true;
  
  try {
    // 测试 1: SSE 连接
    console.log("\n" + "=".repeat(60));
    console.log("测试 1: SSE 连接和消息接收");
    console.log("=".repeat(60));
    
    try {
      const sseResult = await testSSEConnection();
      if (!sseResult.success) {
        console.error("⚠️ SSE 测试未完全通过");
        allTestsPassed = false;
      }
    } catch (error) {
      console.error("❌ SSE 测试失败:", error.message);
      allTestsPassed = false;
    }
    
    // 测试 2: 轮询端点
    console.log("\n" + "=".repeat(60));
    console.log("测试 2: 轮询端点（降级方案）");
    console.log("=".repeat(60));
    
    const pollingResult = await testPollingEndpoint();
    if (!pollingResult.success) {
      allTestsPassed = false;
    }
    
    // 总结
    console.log("\n" + "=".repeat(60));
    console.log("测试总结");
    console.log("=".repeat(60));
    
    if (allTestsPassed) {
      console.log("✅ 所有测试通过！");
      console.log("\n📝 功能状态:");
      console.log("   - SSE 推送: ✅ 可用");
      console.log("   - 轮询降级: ✅ 可用");
      console.log("\n🎉 SSE 优化已成功部署！");
    } else {
      console.log("⚠️ 部分测试失败");
      console.log("\n建议:");
      console.log("   1. 确保服务器正在运行: npm run dev:https");
      console.log("   2. 检查端口 3000 是否被占用");
      console.log("   3. 查看服务器日志排查错误");
    }
    
  } catch (error) {
    console.error("\n❌ 测试执行失败:", error);
    process.exit(1);
  }
}

// 运行测试
main();
