---
name: powerpoint-operations
description: PowerPoint 演示文稿操作技能。用于创建幻灯片、添加内容（文本、图片、形状、表格）、格式设置及生成演示大纲。当用户提及 PPT、幻灯片、演示文稿或演讲时使用。
---

# PowerPoint 操作技能

## 用法

通过生成**隐藏的 Office.js 代码**来操作 PowerPoint，这些代码由前端自动执行，对用户完全透明。

### 重要规则

- **用户体验优先**：用户只需看到自然语言，不应看到任何代码
- **隐藏代码格式**：用 HTML 注释包裹代码：`<!--OFFICE-CODE:powerpoint\ncode\n-->`
- **友好反馈**：操作完成后用自然语言告知用户结果
- **完整且可执行**：生成的代码必须是完整、可直接运行的 Office.js 代码

## ⚠️ 工具选择优先级（强制规则）

### 优先使用 MCP 领域工具

DocuPilot 2.0 提供**领域聚合 MCP 工具**，比通用 execute_code 工具更快、更安全、更易用。

**强制规则**：
1. **默认使用 MCP 领域工具** - 覆盖 85%+ 常见场景
2. **仅在 MCP 工具无法满足需求时使用 execute_code** - 用于复杂高级 API

### 可用 PowerPoint MCP 工具

| 工具 | 用途 | 频率 |
|------|---------|-----------|
| `ppt_shape` | 文本框、图片、形状 | ⭐⭐⭐ 最常用 |
| `ppt_slide` | 幻灯片管理 | ⭐⭐ 常用 |
| `ppt_table` | 表格创建/编辑 | ⭐ 中等 |
| `execute_code` | 动画、主题等 | 备用工具 |

### 工具选择决策树

```
用户请求
  |
  ├─ 添加文本/图片/形状？ → 使用 ppt_shape
  ├─ 管理幻灯片？ → 使用 ppt_slide
  ├─ 创建/编辑表格？ → 使用 ppt_table
  └─ 动画/切换/主题？ → 使用 execute_code
```

### MCP 工具调用方式

```typescript
// ✅ 推荐：使用 MCP 领域工具
mcp__office__ppt_shape({
  action: "add_text",
  slideIndex: 0,
  text: "2024 Annual Report",
  position: { left: 100, top: 150, width: 720, height: 120 }
})

// ❌ 不推荐：除非 MCP 工具无法满足需求
mcp__office__execute_code({
  host: "powerpoint",
  code: "PowerPoint.run(async (context) => { ... })"
})
```

### 示例对比

**场景**：创建带标题幻灯片的演示文稿

**使用 MCP 工具（推荐）**：
```typescript
// 步骤 1：添加新幻灯片
mcp__office__ppt_slide({
  action: "add",
  layout: "Title"
})

// 步骤 2：添加标题文本框
mcp__office__ppt_shape({
  action: "add_text",
  slideIndex: 0,
  text: "2024 Annual Report",
  position: { left: 100, top: 150, width: 720, height: 120 },
  format: {
    fontSize: 60,
    bold: true,
    alignment: "Center",
    fontColor: "#2E5090"
  }
})

// 步骤 3：添加副标题
mcp__office__ppt_shape({
  action: "add_text",
  slideIndex: 0,
  text: "Q4 Financial Summary",
  position: { left: 100, top: 300, width: 720, height: 60 },
  format: {
    fontSize: 32,
    alignment: "Center",
    fontColor: "#4472C4"
  }
})
```

**使用 execute_code（仅必要时）**：
```typescript
// 仅当需要动画或高级功能时
mcp__office__execute_code({
  host: "powerpoint",
  description: "Add slide transition animation",
  code: `
    PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      // Advanced animation configuration...
      await context.sync();
    });
  `
})
```

### 性能对比

| 指标 | MCP 工具 | execute_code | 改进 |
|--------|-----------|--------------|-------------|
| 响应时间 | 1.2s | 2.5s | ↓52% |
| Token 消耗 | ~280 | ~800 | ↓65% |
| 错误率 | <5% | 15% | ↓67% |

### 完整工具 API 参考

详细工具参数和返回值请参考：
- [MCP工具API文档](../../../docs/MCP_TOOLS_API.md)
- [MCP工具完整列表](../../../docs/MCP_TOOLS_REFERENCE.md)

## 工作流

1. **理解需求**：分析用户的演示文稿操作请求
2. **参考模板**：查阅 TOOLS.md 中的代码模板
3. **生成代码**：创建完整的 Office.js 代码
4. **嵌入隐藏标记**：用 `<!--OFFICE-CODE:powerpoint ... -->` 包裹代码
5. **添加友好消息**：告知用户操作结果

## 支持的功能

- **幻灯片管理**：添加、删除、插入（从 Base64）、获取列表。
- **内容添加**：
  - **文本框**：设置文本、字体、颜色、对齐。
  - **形状**：添加几何形状（矩形、圆形等）、线条、连接器。
  - **图片**：插入 Base64 图片，设置位置和大小。
  - **表格**：创建表格、填充数据、设置样式（边框、填充、字体）。
- **格式设置**：调整位置、大小、填充颜色、线条样式。
- **组合/取消组合**：管理形状组。
- **元数据**：添加和读取标签。
- **文本格式增强**：
  - 完整字体属性：粗体、斜体、下划线、字体名、大小、颜色
  - 文本框垂直对齐：顶部、居中、底部
  - 文本框边距设置：左、右、上、下边距控制
  - 文本自动适应：自动缩小以适应形状
- **幻灯片版式和母版**：
  - 获取当前幻灯片的版式和母版信息
  - 使用指定版式创建新幻灯片
  - 列出所有可用幻灯片版式
  - 查询母版名称和 ID
- **主题系统**：
  - 使用主题颜色填充形状（accent1-6、background、text 等）
  - 主题颜色自动适应演示文稿主题
  - 支持 12 种标准主题颜色
- **表格数据操作**：
  - 读取所有表格单元格数据
  - 更新指定单元格内容
  - 设置行高和列宽
  - 创建含合并单元格的表格
  - 格式化表格单元格（背景、字体、边框）

## ⚠️ 常见错误处理

### InvalidArgument 错误
- **原因**：幻灯片索引超出范围、无效的形状 ID
- **解决方案**：使用 `getItemAt()` 时确保索引有效
```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

if (index >= 0 && index < slides.items.length) {
  const slide = slides.getItemAt(index);
  // 操作幻灯片
}
```

### InvalidReference 错误
- **原因**：引用已删除的幻灯片或形状
- **解决方案**：删除后不要继续引用对象

### 文本格式注意
- 所有字体属性在 `textFrame.textRange.font` 对象上
- 垂直对齐在 `textFrame.verticalAlignment`
- 边距单位为磅（points），非像素
- 使用 `FontUnderlineStyle` 枚举设置下划线样式

### 主题颜色用法
- 主题颜色名称：accent1-accent6、background1-background2、text1-text2、hyperlink、followedHyperlink
- 使用主题颜色可确保演示文稿风格一致
- 使用主题颜色的形状在主题变更时会自动更新

### 表格操作限制
- 表格单元格索引从 0 开始
- 读取单元格数据需要多次 sync，注意性能
- 创建表格时通过 `mergedAreas` 参数指定合并单元格
- 表格的 `shape.type` 属性值为 `PowerPoint.ShapeType.table`

### 最佳实践
1. 访问集合前先 load 和 sync
2. 坐标和尺寸使用合理的数值范围
3. 使用标签管理幻灯片元数据
4. 文本框大小应足以容纳内容，避免溢出
5. 使用版式和母版可保持演示文稿风格一致
6. **关键**：使用 `slides.add()` 添加新幻灯片时，必须：
   - 在 `add()` 后立即调用 `await context.sync()`
   - 加载 slides 集合并再次 sync 以获取新幻灯片的引用
   - 然后才能向新幻灯片添加形状
   
   ```javascript
   // ✅ 正确：添加幻灯片、sync、获取引用，再添加形状
   context.presentation.slides.add();
   await context.sync();
   
   const slides = context.presentation.slides;
   slides.load("items");
   await context.sync();
   const newSlide = slides.items[slides.items.length - 1];
   
   // 现在可以安全添加形状
   const textBox = newSlide.shapes.addTextBox("text");
   ```
   
   ```javascript
   // ❌ 错误：add() 后立即使用幻灯片
   const slide = context.presentation.slides.add();
   const textBox = slide.shapes.addTextBox("text"); // 可能静默失败！
   ```

## 示例

### 示例 1：创建标题幻灯片
**用户请求**：「创建标题为「2024 年度报告」的幻灯片」

**你的回复**：
```
我来为你创建标题幻灯片。

[步骤 1：使用 ppt_slide 工具添加新幻灯片]
[步骤 2：使用 ppt_shape 工具添加标题文本框]

✅ 已成功创建标题「2024 Annual Report」的标题幻灯片。
```

**使用的工具**：
1. `mcp__office__ppt_slide({ action: "add", layout: "Title" })`
2. `mcp__office__ppt_shape({ action: "add_text", slideIndex: 0, text: "2024 Annual Report", position: { left: 100, top: 200, width: 720, height: 120 }, format: { fontSize: 60, bold: true, alignment: "Center" } })`

### 示例 2：创建数据表格幻灯片
**用户请求**：「添加带季度销售表格的幻灯片」

**你的回复**：
```
我来创建带销售表格的幻灯片。

[步骤 1：添加新幻灯片]
[步骤 2：使用 ppt_table 工具创建表格]

✅ 已成功创建带季度销售表格的幻灯片。
```

**使用的工具**：
1. `mcp__office__ppt_slide({ action: "add" })`
2. `mcp__office__ppt_table({ action: "create", slideIndex: 1, rows: 4, columns: 5, position: { left: 50, top: 100, width: 700, height: 250 }, data: [["Product", "Q1", "Q2", "Q3", "Q4"], ...], style: { headerRow: true } })`

### 示例 3：复杂幻灯片与自定义格式
**用户请求**：「创建带自定义样式标题和副标题的幻灯片」

**你的回复**：
```
我来为你创建自定义标题幻灯片。

<!--OFFICE-CODE:powerpoint
PowerPoint.run(async (context) => {
  // 添加新幻灯片
  context.presentation.slides.add();
  await context.sync();
  
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  const slide = slides.items[slides.items.length - 1];
  
  // 添加标题文本框
  const textBox = slide.shapes.addTextBox("2024 Annual Report");
  textBox.left = 100;
  textBox.top = 200;
  textBox.width = 760;
  textBox.height = 100;
  
  // 设置标题样式
  textBox.textFrame.textRange.font.size = 60;
  textBox.textFrame.textRange.font.bold = true;
  textBox.textFrame.textRange.paragraphFormat.align = PowerPoint.ParagraphAlignment.center;
  
  await context.sync();
});
-->

✅ 已创建标题「2023 Annual Report」的新幻灯片。
```

### 示例 2：插入数据表格
**用户请求**：「在当前幻灯片插入 3x3 数据表格」

**你的回复**：
```
好的，正在当前幻灯片插入表格。

<!--OFFICE-CODE:powerpoint
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const table = slide.shapes.addTable(3, 3, {
    left: 150,
    top: 150,
    width: 600,
    height: 300,
    values: [
      ["Quarter", "Revenue", "Expense"],
      ["Q1", "1M", "800K"],
      ["Q2", "1.2M", "900K"]
    ],
    uniformCellProperties: {
      fill: { color: "#4472C4" },
      font: { color: "white", size: 18 }
    }
  });
  
  await context.sync();
});
-->

✅ 已插入 3x3 数据表格并应用蓝色样式。
```

## 用户文件处理

### 工作区文件访问

用户可将文件（如图片）上传到当前会话，这些文件保存在工作区目录：

- **用户上传文件路径**：`workspace/sessions/{session_id}/uploads/`
- **生成文件保存路径**：`workspace/sessions/{session_id}/outputs/`

### 文件操作流程

1. **查找用户上传的图片**：
   ```typescript
   // 使用 Glob 工具查找图片文件
   // 文件名包含时间戳前缀，使用通配符
   const pattern = "workspace/sessions/{session_id}/uploads/*.{png,jpg,jpeg}";
   ```

2. **处理图片**：
   - PowerPoint 图片插入需要 Base64 编码
   - 若用户上传图片，需先读取文件并转换为 Base64

3. **保存演示大纲**：
   ```typescript
   // 使用 Write 工具保存演示大纲
   Write: workspace/sessions/{session_id}/outputs/presentation_outline.txt
   ```

### 示例工作流

**用户请求**：「用我上传的图片创建演示文稿」

**处理步骤**：
1. 使用 Glob 查找：`workspace/sessions/abc123/uploads/*.png`
2. 告知用户：「我找到你上传的图片 `logo.png`。我会创建演示文稿并插入这张图片。」
3. 读取图片文件，必要时转换为 Base64
4. 使用 Office.js API 创建幻灯片并插入图片
5. 如需保存演示大纲，保存到：`workspace/sessions/abc123/outputs/slides_outline.txt`

## 详细模板

更多操作模板请参考 [TOOLS.md](TOOLS.md)。
