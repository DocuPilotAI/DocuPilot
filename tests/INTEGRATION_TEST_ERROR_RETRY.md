# 错误自我修正架构 - 集成测试指南

## 测试目标

验证 Agent 能够自动捕获代码执行错误，并根据错误反馈重新生成正确的代码。

## 测试环境要求

1. **Office 应用**: Excel（Web 或桌面版）
2. **浏览器**: Chrome/Edge（最新版本）
3. **DocuPilot**: 已启动开发服务器（`npm run dev`）

## 测试场景

### 场景 1: InvalidReference 错误 - 工作表不存在

**测试步骤**:

1. 在 Excel 中打开一个新工作簿（默认只有 Sheet1）
2. 在 DocuPilot 中输入: `请在 Sheet2 中写入 "Hello World"`
3. 观察自动重试过程

**预期结果**:

- ❌ 第一次尝试：代码生成使用 `getItem("Sheet2")`，执行失败（Sheet2 不存在）
- 🔄 自动重试：系统显示"正在自动重试（1/3）..."
- ✅ 第二次尝试：Agent 根据错误反馈生成修正代码，使用 `getItemOrNullObject()` 或创建新工作表
- ✅ 成功执行：操作完成

**验证点**:

```typescript
// 第一次生成的代码（可能失败）
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItem("Sheet2");
  // ...
});

// 修正后的代码（应该成功）
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject("Sheet2");
  sheet.load("name");
  await context.sync();
  
  if (sheet.isNullObject) {
    const newSheet = context.workbook.worksheets.add("Sheet2");
    newSheet.activate();
  }
  // ...
});
```

### 场景 2: InvalidArgument 错误 - 参数格式错误

**测试步骤**:

1. 在 Excel 中打开一个工作簿
2. 在 DocuPilot 中输入: `请在单元格 ABC123 中写入数字 100`（故意使用无效地址）
3. 观察自动重试过程

**预期结果**:

- ❌ 第一次尝试：使用无效的单元格地址，执行失败
- 🔄 自动重试：系统检测到 InvalidArgument 错误
- ✅ 第二次尝试：Agent 修正地址格式或使用有效地址
- ✅ 成功执行

### 场景 3: 复杂操作 - 多次重试

**测试步骤**:

1. 在 Excel 中打开工作簿
2. 输入复杂请求: `在不存在的工作表 "数据分析" 中创建一个表格，包含姓名、年龄、城市三列，并填充5行示例数据`
3. 观察多次重试过程

**预期结果**:

- 可能需要 2-3 次重试
- 每次重试都有明确的错误提示和改进
- 最终成功创建表格或给出清晰的失败原因

## 手动测试步骤

### 步骤 1: 启动应用

```bash
cd /Users/cw/Documents/GitHub/github/DocuPilot
npm run dev
```

### 步骤 2: 在 Excel 中加载插件

1. 打开 Excel（Web 版推荐: https://www.office.com/launch/excel）
2. 插入 -> Office 加载项 -> 上传我的加载项
3. 选择 `manifest.xml`
4. 加载 DocuPilot 任务窗格

### 步骤 3: 执行测试用例

使用上述场景中的测试输入，逐个验证。

### 步骤 4: 检查控制台日志

打开浏览器开发者工具（F12），查看：

```javascript
// 应该看到类似的日志
[useChat] Office code execution failed: {...}
[useChat] Auto-retrying (1/3)...
[useChat] SSE event message ...
```

### 步骤 5: 验证 UI 显示

在聊天界面中，错误消息应该显示：

- 重试次数指示器（🔄 重试 1/3）
- 错误类型标签（如 "InvalidReference"）
- 清晰的错误说明

## 自动化测试（可选）

如果需要更系统的测试，可以使用现有的测试框架：

```bash
# 运行 Office Skills 测试套件
npm run test:office-skills
```

测试文件位置:
- `tests/office-skills/test-cases/excel-test-cases.json`
- `tests/office-skills/test-runner/TestExecutor.ts`

## 性能基准

**预期指标**:

- 单次重试响应时间: 3-5 秒
- 最大总时间（3次重试）: < 15 秒
- 成功率提升: 从 ~60% 提升到 80-85%

**实际测量方法**:

1. 在浏览器控制台中运行：

```javascript
// 开始计时
console.time('retry-test');

// 执行测试用例（触发错误）
// ...

// 结束计时
console.timeEnd('retry-test');
```

2. 统计成功率：

```javascript
// 运行 10 个测试用例，记录成功/失败次数
let success = 0;
let failed = 0;

// ... 执行测试 ...

console.log(`成功率: ${(success / (success + failed) * 100).toFixed(1)}%`);
```

## 常见问题排查

### 问题 1: 重试未触发

**可能原因**:
- 错误未被正确捕获
- `retryStateRef` 未正确初始化

**检查方法**:
```javascript
// 在 use-chat.ts 中添加日志
console.log('[Debug] Retry state:', retryStateRef.current);
console.log('[Debug] Current retries:', currentRetries);
```

### 问题 2: 重试次数超限

**可能原因**:
- `MAX_RETRIES` 常量未生效
- 重试计数未正确更新

**检查方法**:
```javascript
// 验证 MAX_RETRIES
console.log('[Debug] MAX_RETRIES:', MAX_RETRIES);
```

### 问题 3: Agent 生成的修正代码仍然错误

**可能原因**:
- 错误反馈信息不够详细
- Agent 系统提示需要优化

**解决方案**:
- 查看 `buildErrorFeedback()` 生成的反馈内容
- 调整 `chat/route.ts` 中的错误处理提示

## 测试报告模板

```markdown
## 测试执行记录

**日期**: YYYY-MM-DD
**测试人**: 名字
**环境**: Excel Web / Excel Desktop

### 场景 1: InvalidReference 测试
- [ ] 第一次尝试失败 ✅
- [ ] 自动重试触发 ✅
- [ ] 第二次尝试成功 ✅
- [ ] UI 显示正确 ✅
- [ ] 响应时间: ___ 秒

### 场景 2: InvalidArgument 测试
- [ ] 第一次尝试失败 ✅
- [ ] 自动重试触发 ✅
- [ ] 第二次尝试成功 ✅
- [ ] UI 显示正确 ✅
- [ ] 响应时间: ___ 秒

### 场景 3: 复杂操作测试
- [ ] 需要多次重试 ✅
- [ ] 最终成功 ✅
- [ ] 总响应时间: ___ 秒

### 总体评估
- 成功率: ____%
- 平均重试次数: ___
- 用户体验: 好 / 中 / 差
- 需要改进的地方: ___
```

## 后续优化建议

基于测试结果，可以考虑以下优化：

1. **错误模式库扩展**: 根据实际遇到的错误，补充 `error-patterns.ts`
2. **提示优化**: 根据 Agent 的修正效果，调整 `chat/route.ts` 中的提示
3. **性能优化**: 如果重试时间过长，考虑缓存常见错误的修复方案
4. **用户体验**: 添加取消重试的按钮，或调整重试策略

## 完成标准

测试通过的标准：

✅ 所有三个场景都能成功触发自动重试  
✅ 至少 80% 的错误能在 2 次重试内解决  
✅ UI 正确显示重试状态和错误信息  
✅ 控制台日志清晰，便于调试  
✅ 响应时间在可接受范围内（< 15 秒）  
✅ 代码无 linter 错误
