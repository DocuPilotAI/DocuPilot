# SSE 推送功能测试指南

## 测试目的

验证 SSE（Server-Sent Events）任务推送功能是否正常工作，包括：
1. SSE 连接建立
2. 任务即时推送
3. 自动重连
4. 降级到轮询

## 前置条件

确保服务器正在运行：
```bash
npm run dev:https
```

验证服务器状态：
```bash
ps aux | grep "node.*server.mjs" | grep -v grep
# 应该看到 server.mjs 进程
```

## 测试步骤

### 测试 1: 浏览器开发工具测试

1. **打开浏览器开发工具**
   - 打开 Chrome/Edge 浏览器
   - 访问 `https://localhost:3000/taskpane`
   - 按 F12 打开开发者工具

2. **打开 Network 面板**
   - 切换到 Network 标签
   - 筛选：`task-stream` 或 `EventStream`

3. **发送 Office 操作请求**
   - 在聊天框输入："在 A1 单元格写入 Hello"
   - 点击发送

4. **观察 Network 请求**

**预期结果（SSE 模式）**：
```
✅ 应该看到：
   - 1 个 task-stream 连接（Type: eventsource）
   - 状态：pending（保持连接）
   - 无 tool-result?action=pending_executions 轮询请求

❌ 不应该看到：
   - 大量的 GET tool-result?action=pending_executions 请求
```

**预期结果（降级到轮询）**：
```
如果 SSE 失败（如网络问题），会看到：
   - 控制台警告："falling back to polling"
   - 大量的 GET tool-result?action=pending_executions 请求（每 200ms）
```

5. **观察 Console 日志**

**SSE 模式日志**：
```
[useChat/SSE] Connected successfully
[useChat/SSE] Connection established
[useChat/SSE] Task received: <correlationId>
[useChat] MCP Tool task received: <correlationId> <host>
```

**降级模式日志**：
```
[useChat/SSE] Max reconnect attempts (3) reached, falling back to polling
[useChat/Polling] Started (SSE unavailable)
[useChat] MCP Tool task received: <correlationId> <host>
```

### 测试 2: SSE 端点直接测试

使用浏览器直接访问 SSE 端点：

```javascript
// 在浏览器控制台执行
const eventSource = new EventSource('https://localhost:3000/api/task-stream');

eventSource.onopen = () => {
  console.log('✅ SSE 连接成功');
};

eventSource.onmessage = (event) => {
  console.log('📨 收到消息:', event.data);
  const data = JSON.parse(event.data);
  if (data.type === 'connected') {
    console.log('✅ 连接确认消息');
  } else if (data.type === 'task') {
    console.log('📦 任务推送:', data.correlationId, data.host);
  }
};

eventSource.onerror = (error) => {
  console.error('❌ SSE 错误:', error);
};

// 测试完成后关闭
// eventSource.close();
```

**预期输出**：
```
✅ SSE 连接成功
📨 收到消息: {"type":"connected","timestamp":1234567890}
✅ 连接确认消息
```

### 测试 3: 性能对比测试

#### 测试 SSE 模式

1. 打开 Network 面板并清空
2. 发送 10 个 Office 操作请求
3. 统计 Network 请求数量

**预期结果**：
- task-stream 连接：1 个（保持连接）
- tool-result POST 请求：10 个（提交结果）
- **总计：11 个请求**

#### 测试轮询模式

为了对比，临时禁用 SSE：

1. 修改 `src/lib/use-chat.ts`：
   ```typescript
   const MCP_SSE_ENABLED = false; // 临时禁用
   ```

2. 重启服务器

3. 重复测试步骤

**预期结果**：
- tool-result GET 请求：~300 个（每 200ms 轮询，假设 1 分钟完成）
- tool-result POST 请求：10 个
- **总计：~310 个请求**

**性能提升**：
- 请求数减少：(310 - 11) / 310 = **96.5%**
- 网络开销降低：约 **30 倍**

### 测试 4: 自动重连测试

1. 建立 SSE 连接
2. 在服务器终端按 Ctrl+C 停止服务器
3. 观察浏览器控制台

**预期行为**：
```
[useChat/SSE] Connection error: ...
[useChat/SSE] Reconnecting... (attempt 1/3)
[useChat/SSE] Reconnecting... (attempt 2/3)
[useChat/SSE] Reconnecting... (attempt 3/3)
[useChat/SSE] Max reconnect attempts (3) reached, falling back to polling
[useChat/Polling] Started (SSE unavailable)
```

4. 重启服务器
5. 刷新页面，验证恢复使用 SSE

### 测试 5: 多任务并发推送

1. 准备一个复杂的请求，会产生多个 MCP Tool 调用
2. 观察 Console 和 Network

**预期行为**：
- 每个任务都能即时推送
- 无任务丢失
- 执行顺序正确

## 性能指标

### SSE 模式

| 指标 | 目标值 | 验证方法 |
|------|--------|----------|
| 任务推送延迟 | < 1ms | Console 时间戳对比 |
| HTTP 请求数 | 1 个长连接 + N 个 POST | Network 面板统计 |
| 连接稳定性 | 保持连接 | 观察连接状态 |
| 心跳间隔 | 30 秒 | 等待观察心跳消息 |

### 降级到轮询

| 指标 | 目标值 | 验证方法 |
|------|--------|----------|
| 降级触发 | 3 次重连失败后 | Console 日志 |
| 轮询间隔 | 200ms | Network 请求间隔 |
| 功能正常 | 任务仍能执行 | 实际操作测试 |

## 故障排查

### 问题 1: SSE 连接失败

**可能原因**：
- 服务器未启动
- 端口被占用
- 证书问题

**解决方法**：
```bash
# 检查服务器状态
ps aux | grep server.mjs

# 检查端口
lsof -i :3000

# 重启服务器
npm run dev:https
```

### 问题 2: SSE 频繁断开

**可能原因**：
- 网络不稳定
- 代理/防火墙限制
- 浏览器限制

**解决方法**：
- 查看 Console 错误日志
- 检查网络配置
- 会自动降级到轮询，功能不受影响

### 问题 3: 任务没有推送

**可能原因**：
- EventEmitter 未触发
- SSE 连接未建立
- 任务已被处理

**解决方法**：
```bash
# 查看服务器日志
# 应该看到：
# [SSE] New client connection established
# [SSE] Pushed task <id> (<host>) to client
```

## 成功标准

所有以下条件都满足：

- ✅ SSE 连接能成功建立
- ✅ 任务能即时推送到前端
- ✅ Office 操作能正常执行
- ✅ Network 面板无轮询请求（SSE 模式）
- ✅ 连接断开后能自动重连
- ✅ 多次重连失败后能降级到轮询
- ✅ 降级后功能仍正常

## 性能对比数据

测试完成后，记录以下数据：

| 测试场景 | SSE 模式 | 轮询模式 | 改善 |
|---------|----------|----------|------|
| 1 分钟内 HTTP 请求数 | | | |
| 任务推送延迟（平均） | | | |
| 网络流量消耗 | | | |
| 浏览器 CPU 占用 | | | |

## 总结

SSE 优化应该带来：
- 请求数减少 96%+
- 任务推送延迟 < 1ms
- 更流畅的用户体验
- 保持功能完整性（自动降级）

---

**测试日期**：2026-02-03  
**测试版本**：v1.0.0（SSE 优化）  
**测试状态**：等待浏览器测试验证
