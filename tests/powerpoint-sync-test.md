# PowerPoint Sync 测试用例

## 测试目的
验证修复后的PowerPoint代码生成逻辑，确保在添加新幻灯片后正确使用 `context.sync()` 并能成功添加shapes。

## 测试环境要求
- PowerPoint Online 或 PowerPoint 桌面版 (Microsoft 365)
- DocuPilot Add-in 已加载
- 开启浏览器开发者工具 Console 查看日志

## 测试用例 1: 添加新幻灯片并插入文本框

### 输入提示词
```
在新幻灯片中添加一个标题文本框，内容为"测试幻灯片"
```

### 期望行为
1. AI生成代码并执行
2. PowerPoint中出现新的幻灯片
3. 新幻灯片中包含一个文本框，显示"测试幻灯片"
4. 控制台日志显示执行成功

### 期望生成的代码结构
```javascript
PowerPoint.run(async (context) => {
  // 添加新幻灯片
  context.presentation.slides.add();
  await context.sync(); // 关键步骤！
  
  // 获取新添加的幻灯片
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  const newSlide = slides.items[slides.items.length - 1];
  
  // 添加文本框
  const textBox = newSlide.shapes.addTextBox("测试幻灯片");
  textBox.left = 100;
  textBox.top = 100;
  textBox.width = 300;
  textBox.height = 50;
  
  await context.sync();
});
```

### 验证点
- [ ] 新幻灯片已创建
- [ ] 文本框已添加并显示正确内容
- [ ] Console 中显示 `✅ Execution successful`
- [ ] 执行时间 > 100ms (说明确实执行了异步操作)
- [ ] 没有错误日志

---

## 测试用例 2: 创建复杂幻灯片（多个shapes）

### 输入提示词
```
创建一个包含标题、矩形和文本框的幻灯片
```

### 期望行为
1. 创建新幻灯片
2. 添加三个shapes：
   - 标题文本框
   - 一个蓝色矩形
   - 一个描述文本框

### 期望生成的代码结构
```javascript
PowerPoint.run(async (context) => {
  context.presentation.slides.add();
  await context.sync();
  
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  const newSlide = slides.items[slides.items.length - 1];
  
  // 添加标题
  const title = newSlide.shapes.addTextBox("标题");
  title.left = 50;
  title.top = 50;
  title.width = 600;
  title.height = 80;
  title.textFrame.textRange.font.size = 40;
  title.textFrame.textRange.font.bold = true;
  
  // 添加矩形
  const rect = newSlide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect.left = 50;
  rect.top = 150;
  rect.width = 200;
  rect.height = 100;
  rect.fill.setSolidColor("blue");
  
  // 添加文本框
  const desc = newSlide.shapes.addTextBox("描述文本");
  desc.left = 300;
  desc.top = 150;
  desc.width = 300;
  desc.height = 100;
  
  await context.sync();
});
```

### 验证点
- [ ] 新幻灯片包含3个shapes
- [ ] 标题文本粗体且字号为40
- [ ] 矩形为蓝色
- [ ] 所有shapes位置正确
- [ ] Console 无错误

---

## 测试用例 3: 使用选中的幻灯片（无需add）

### 前置条件
手动在PowerPoint中选中一个空白幻灯片

### 输入提示词
```
在当前幻灯片添加一个圆形
```

### 期望行为
不创建新幻灯片，在选中的幻灯片上添加圆形

### 期望生成的代码结构
```javascript
PowerPoint.run(async (context) => {
  // 获取选中的幻灯片，无需add()
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 直接添加圆形
  const circle = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.oval);
  circle.left = 200;
  circle.top = 200;
  circle.width = 150;
  circle.height = 150;
  circle.fill.setSolidColor("green");
  
  await context.sync();
});
```

### 验证点
- [ ] 没有创建新幻灯片
- [ ] 选中的幻灯片上有绿色圆形
- [ ] 圆形尺寸和位置正确

---

## 测试用例 4: Transformer架构图（复杂场景）

### 输入提示词
```
Draw a diagram of the Transformer architecture and explain its components in the powerpoint.
```

### 期望行为
1. 创建多张幻灯片
2. 第一张：Transformer架构总览图
3. 包含多个文本框和形状
4. 所有元素正确显示

### 验证点
- [ ] 至少创建1张幻灯片
- [ ] 包含多个形状（矩形、文本框等）
- [ ] 形状有不同颜色和位置
- [ ] Console 显示执行成功
- [ ] 执行时间合理（300-2000ms）

---

## 调试步骤

### 如果测试失败：

1. **检查 Console 日志**
   - 查找 `[Office Code Executor]` 开头的日志
   - 查找 `[useChat] MCP Tool` 开头的日志
   - 查找 `[MCP/office]` 开头的日志

2. **启用 DEBUG 模式**
   在 `.env.local` 文件中添加：
   ```
   NEXT_PUBLIC_DEBUG_CHAT=1
   ```
   重启开发服务器

3. **检查关键日志**
   - `Code length:` - 代码长度
   - `Execution successful` - 执行是否成功
   - `Duration: Xms` - 执行时间
   - `Result:` - 返回结果

4. **常见问题诊断**

   **问题**: 显示"执行成功"但PPT没有内容
   - 检查执行时间是否 < 100ms
   - 查看是否有 "possible silent failure" 警告
   - 原因：代码缺少必要的 `context.sync()`

   **问题**: InvalidArgument 错误
   - 检查是否在 `slides.add()` 后立即使用返回的slide对象
   - 确认使用了正确的API调用顺序

   **问题**: 超时
   - 检查PowerPoint是否正常运行
   - 查看网络连接
   - 确认前端轮询正常工作

---

## 自动化测试脚本

可以使用以下Node.js脚本进行自动化测试（需要在PowerPoint环境中运行）：

```javascript
// test-powerpoint-sync.js
// 在浏览器Console中运行

async function testCase1() {
  console.log('=== Test Case 1: Add Slide with TextBox ===');
  
  const code = `
    PowerPoint.run(async (context) => {
      context.presentation.slides.add();
      await context.sync();
      
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();
      
      const newSlide = slides.items[slides.items.length - 1];
      
      const textBox = newSlide.shapes.addTextBox("测试幻灯片");
      textBox.left = 100;
      textBox.top = 100;
      textBox.width = 300;
      textBox.height = 50;
      
      await context.sync();
      return "Success";
    });
  `;
  
  const result = await executeOfficeCode('powerpoint', code);
  console.log('Test Case 1 Result:', result);
  return result.success;
}

// 运行测试
testCase1().then(success => {
  console.log('Test Case 1:', success ? '✅ PASSED' : '❌ FAILED');
});
```

---

## 回归测试清单

在每次修改代码后，运行以下测试确保没有破坏现有功能：

- [ ] 测试用例 1: 基本文本框添加
- [ ] 测试用例 2: 多个shapes添加
- [ ] 测试用例 3: 使用选中的幻灯片
- [ ] 测试用例 4: Transformer架构图
- [ ] Excel 基本操作仍然正常
- [ ] Word 基本操作仍然正常

---

## 性能基准

| 操作 | 预期时间 | 警告阈值 |
|------|---------|---------|
| 添加单个幻灯片 | 150-300ms | > 1000ms |
| 添加10个shapes | 300-800ms | > 2000ms |
| 复杂图表生成 | 500-2000ms | > 5000ms |

---

## 已知限制

1. PowerPoint Online 某些API可能不可用
2. 大量shapes（>50个）可能导致性能问题
3. 某些形状类型在不同Office版本中支持不同
