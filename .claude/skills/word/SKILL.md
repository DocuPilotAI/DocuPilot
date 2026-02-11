---
name: word-operations
description: Word 文档操作技能。用于读取文档、插入文本/图片/表格、格式化内容、内容控件、页眉/页脚及文档搜索。当用户提及 Word、文档、段落、表格或文字处理时使用。
---

# Word 操作技能

## 用法

通过生成**隐藏的 Office.js 代码**来操作 Word，这些代码由前端自动执行，对用户完全透明。

### 重要规则

- **用户体验优先**：用户只需看到自然语言，不应看到任何代码
- **隐藏代码格式**：用 HTML 注释包裹代码：`<!--OFFICE-CODE:word\ncode\n-->`
- **友好反馈**：操作完成后用自然语言告知用户结果
- **完整且可执行**：生成的代码必须是完整、可直接运行的 Office.js 代码

## ⚠️ 工具选择优先级（强制规则）

### 优先使用 MCP 领域工具

DocuPilot 2.0 提供**领域聚合 MCP 工具**，比通用 execute_code 工具更快、更安全、更易用。

**强制规则**：
1. **默认使用 MCP 领域工具** - 覆盖 85%+ 常见场景
2. **仅在 MCP 工具无法满足需求时使用 execute_code** - 用于复杂高级 API

### 可用 Word MCP 工具

| 工具 | 用途 | 频率 |
|------|---------|-----------|
| `word_paragraph` | 段落插入、格式设置 | ⭐⭐⭐ 最常用 |
| `word_document` | 文档读取、搜索/替换 | ⭐⭐ 常用 |
| `word_table` | 表格创建/编辑 | ⭐ 中等 |
| `execute_code` | 域、批注等 | 备用工具 |

### 工具选择决策树

```
用户请求
  |
  ├─ 插入/格式化段落？ → 使用 word_paragraph
  ├─ 读取/搜索/替换文档？ → 使用 word_document
  ├─ 创建/编辑表格？ → 使用 word_table
  └─ 域/批注/页眉/页脚？ → 使用 execute_code
```

### MCP Tool Invocation Method

```typescript
// ✅ 推荐：使用 MCP 领域工具
mcp__office__word_paragraph({
  action: "insert",
  text: "Chapter 1: Introduction",
  location: "End",
  format: { style: "Heading 1" }
})

// ❌ 不推荐：除非 MCP 工具无法满足需求
mcp__office__execute_code({
  host: "word",
  code: "Word.run(async (context) => { ... })"
})
```

### 示例对比

**场景**：创建报告文档

**使用 MCP 工具（推荐）**：
```typescript
// 步骤 1：插入标题
mcp__office__word_paragraph({
  action: "insert",
  text: "Financial Analysis Report",
  location: "End",
  format: {
    style: "Heading 1",
    alignment: "Center",
    font: { size: 28, bold: true }
  }
})

// 步骤 2：插入正文
mcp__office__word_paragraph({
  action: "insert",
  text: "Chapter 1: Overview\n\nThis report analyzes...",
  location: "End",
  format: {
    style: "Normal",
    font: { size: 12 }
  }
})

// 步骤 3：插入表格
mcp__office__word_table({
  action: "create",
  rows: 3,
  columns: 4,
  data: [
    ["Item", "Q1", "Q2", "Q3"],
    ["Revenue", "$1M", "$1.2M", "$1.5M"],
    ["Cost", "$600K", "$700K", "$800K"]
  ],
  location: "End",
  style: "Grid Table 4 - Accent 1"
})
```

**使用 execute_code（仅必要时）**：
```typescript
// 仅当需要域（如动态日期、目录）时
mcp__office__execute_code({
  host: "word",
  description: "Insert auto-updating date field",
  code: `
    Word.run(async (context) => {
      const range = context.document.body.getRange("End");
      range.insertField(
        Word.InsertLocation.end,
        Word.FieldType.date,
        '\\\\@ "yyyy-MM-dd"',
        true
      );
      await context.sync();
    });
  `
})
```

### 性能对比

| 指标 | MCP 工具 | execute_code | 改进 |
|--------|-----------|--------------|-------------|
| Response Time | 1.2s | 2.5s | ↓52% |
| Token Cost | ~280 | ~800 | ↓65% |
| Error Rate | <5% | 15% | ↓67% |

### 完整工具 API 参考

详细工具参数和返回值请参考：
- [MCP Tools API Documentation](../../../docs/MCP_TOOLS_API.md)
- [MCP Tools Complete List](../../../docs/MCP_TOOLS_REFERENCE.md)

## 工作流

1. **理解需求**：分析用户的操作请求
2. **参考模板**：查阅 TOOLS.md 中的代码模板
3. **生成代码**：创建完整的 Office.js 代码
4. **嵌入隐藏标记**：用 `<!--OFFICE-CODE:word ... -->` 包裹代码
5. **添加友好消息**：告知用户操作结果

## 支持的功能

- **文档编辑**：插入文本、段落、列表、表格、图片（Base64）。
- **内容控件**：创建、读取、更新内容控件（用于表单/模板）。
- **页眉和页脚**：修改文档页眉和页脚内容。
- **格式设置**：设置字体、颜色、段落间距、对齐。
- **搜索与替换**：常规搜索、通配符搜索、批量替换、高亮。
- **文档结构**：操作节、段落、标题样式。
- **域**：
  - 日期域：插入自动更新的当前日期
  - 目录域（TOC）：自动生成文档目录
  - 超链接域：创建可点击链接
  - 页码域：插入页码
  - Addin 域：存储自定义插件数据
  - 更新域：刷新域内容
- **脚注和尾注**：
  - 在文档中插入脚注引用
  - 插入尾注引用
  - 读取和管理脚注/尾注内容
- **样式管理**：
  - 应用内置样式（Heading1-9、Body、Quote 等）
  - 使用 `styleBuiltIn` 枚举应用样式
  - 批量应用样式到匹配内容
  - 获取和复制现有段落样式
- **批注**：
  - 插入语法和写作建议批注（需要 Microsoft 365 订阅）
  - 向段落添加彩色下划线提示
  - 提供弹出建议和操作
  - 注册和响应批注事件（点击、悬停等）
  - 删除批注

## ⚠️ 常见错误处理

### InvalidArgument 错误
- **原因**：引用的内容控件、段落或区域不存在
- **解决方案**：操作前检查选择或内容是否存在
```javascript
const selection = context.document.getSelection();
selection.load("text");
await context.sync();
if (selection.text.length === 0) {
  console.log("Nothing selected");
  return;
}
```

### InvalidReference 错误
- **原因**：访问已删除的对象
- **解决方案**：删除后不要继续引用对象

### 域相关错误
- **问题**：域代码格式不正确
- **解决方案**：使用正确的域开关格式，如日期域 `\\@ "M/d/yyyy"`
- **问题**：域不更新
- **解决方案**：调用 `field.update()` 方法手动更新域

### 批注相关限制
- 批注 API 需要 Microsoft 365 订阅，不支持一次性购买的 Office 版本
- 批注不会持久保存到文档，重新打开文档后需重新生成
- 若用户接受建议并保存文档，更改会保留

### 最佳实践
1. 读取属性前调用 `load()` 和 `context.sync()`
2. 避免在循环中频繁调用 `context.sync()`
3. 使用批量操作以提升性能
4. 域操作后调用 `load()` 获取结果和代码
5. 插入脚注/尾注后可加载 `reference` 属性获取引用编号

## 示例

### 示例 1：插入多个段落
**用户请求**：「添加标题为「执行摘要」的章节及导言」

**你的回复**：
```
I'll add the chapter heading and introduction for you.

[Step 1: Inserting heading using word_paragraph tool]
[Step 2: Inserting body text using word_paragraph tool]

✅ Successfully added chapter "Executive Summary" with heading and introductory paragraph.
```

**Tools Used**:
1. `mcp__office__word_paragraph({ action: "insert", text: "Executive Summary", location: "End", format: { style: "Heading 1" } })`
2. `mcp__office__word_paragraph({ action: "insert", text: "This report provides...", location: "End", format: { style: "Normal" } })`

### 示例 2：创建数据表格
**用户请求**：「创建显示 Q1-Q4 销售数据的表格」

**你的回复**：
```
I'll create a sales table for you.

[Creating table using word_table tool]

✅ Successfully created a 5x5 sales table with formatted headers.
```

**Tool Used**: 
```typescript
mcp__office__word_table({
  action: "create",
  rows: 5,
  columns: 5,
  data: [
    ["Product", "Q1", "Q2", "Q3", "Q4"],
    ["Product A", "$100K", "$120K", "$150K", "$180K"],
    ["Product B", "$80K", "$90K", "$95K", "$100K"],
    ["Product C", "$60K", "$70K", "$80K", "$90K"],
    ["Total", "$240K", "$280K", "$325K", "$370K"]
  ],
  location: "End",
  style: "Grid Table 4 - Accent 1"
})
```

### 示例 3：插入诗歌（复杂格式）
**用户请求**：「在文档末尾添加李白的《静夜思》」

**你的回复**：
```
I'll insert the poem with proper formatting.

<!--OFFICE-CODE:word
Word.run(async (context) => {
  const body = context.document.body;
  
  // Insert title
  const title = body.insertParagraph("Quiet Night Thoughts", "End");
  title.font.size = 16;
  title.font.bold = true;
  title.alignment = Word.Alignment.centered;
  
  // Insert author
  const author = body.insertParagraph("Li Bai", "End");
  author.font.size = 12;
  author.alignment = Word.Alignment.centered;
  author.spaceAfter = 12;
  
  // Insert poem lines
  const lines = ["Moonlight before my bed,", "I thought it was frost on the ground.", "Lifting my head, I gaze at the bright moon,", "Lowering my head, I think of my hometown."];
  for (const line of lines) {
    body.insertParagraph(line, "End").alignment = Word.Alignment.centered;
  }
  
  await context.sync();
});
-->

✅ Successfully inserted 'Quiet Night Thoughts', including title, author, and complete poem lines, set to center alignment.
```

**说明**：带循环的复杂多元素格式使用 execute_code 更高效。

### 示例 4：创建内容控件表格（高级 API）
**用户请求**：「创建带内容控件的客户信息表格」

**你的回复**：
```
Alright, creating customer information table template for you.

<!--OFFICE-CODE:word
Word.run(async (context) => {
  const body = context.document.body;
  const table = body.insertTable(2, 2, "End", [["Customer Name", ""], ["Contact", ""]]);
  
  // Add content controls
  const nameCell = table.getCell(0, 1);
  const nameCC = nameCell.body.insertContentControl();
  nameCC.title = "Customer Name";
  nameCC.tag = "CustomerName";
  nameCC.placeholderText = "Please enter customer name";
  
  const contactCell = table.getCell(1, 1);
  const contactCC = contactCell.body.insertContentControl();
  contactCC.title = "Contact";
  contactCC.tag = "ContactInfo";
  contactCC.placeholderText = "Please enter phone or email";
  
  await context.sync();
});
-->

✅ Created table with content controls, can be used to fill customer information.
```

## 用户文件处理

### 工作区文件访问

用户可将文件上传到当前会话，这些文件保存在工作区目录：

- **用户上传文件路径**：`workspace/sessions/{session_id}/uploads/`
- **生成文件保存路径**：`workspace/sessions/{session_id}/outputs/`

### 文件操作流程

1. **查找用户上传的文档**：
   ```typescript
   // 使用 Glob 工具查找 Word 文档
   // 文件名包含时间戳前缀，使用通配符
   const pattern = "workspace/sessions/{session_id}/uploads/*.docx";
   ```

2. **读取文本文件**：
   - 纯文本文件（TXT、MD）使用 Read 工具直接读取内容
   - Word 文档引导用户用 Word 打开后使用 Office.js API 操作

3. **保存处理结果**：
   ```typescript
   // 使用 Write 工具保存处理后的文本
   Write: workspace/sessions/{session_id}/outputs/formatted_text.txt
   ```

### 示例工作流

**用户请求**：「帮我格式化上传的文档」

**处理步骤**：
1. Use Glob to find: `workspace/sessions/abc123/uploads/*.docx`
2. Guide user: "I found your uploaded document `report.docx`. Please open this file in Word, then I can help you format it."
3. After user opens file in Word, use Office.js API to apply formatting
4. If need to save processed text version, save to: `workspace/sessions/abc123/outputs/formatted_report.txt`

## 🚨 分步执行规则（强制 / Step-by-Step Execution Rules）

### 核心原则

**复杂任务必须分步执行**，禁止一次性生成超过 30 行或包含超过 5 个主要操作的代码。

### 复杂度限制

| 限制项 | 阈值 | 说明 |
|--------|------|------|
| 代码行数 | ≤ 30 行 | 超过需拆分 |
| insert* 操作数 | ≤ 5 个 | 每步最多 5 个插入操作 |
| 章节数 | 1 个 | 每步只创建 1 个章节 |

### 分步执行流程

对于复杂任务（如创建完整报告模板），必须：

1. **第一步：创建封面/标题**
   - 只创建文档标题和基本信息
   - 返回验证结果

2. **第二步：创建第一个章节**
   - 添加章节标题和内容
   - 返回已创建的段落数

3. **第三步～第N步：依次创建后续章节**
   - 每步只处理一个章节
   - 每步都验证结果

4. **最后一步：添加页眉页脚（如需要）**

### 验证机制

每次执行代码必须返回验证信息：

```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  // 执行操作...
  const title = body.insertParagraph("章节标题", "End");
  title.style = "Heading 1";
  
  await context.sync();
  
  // 必须返回验证信息
  return {
    success: true,
    created: "1个标题段落",
    preview: "章节标题"
  };
});
```

### 禁止的操作

以下操作在分步执行中**禁止使用**：

1. **`body.clear()`** - 会清空整个文档
2. **`insertParagraph(..., "Start")`** - 在开头插入会打乱结构
3. **复杂的 `search()` 定位** - 依赖前面步骤的内容可能找不到
4. **`insertField()` 用于目录** - API 不稳定，容易失败
5. **单次超过 5 个 `insertBreak()`** - 分页符过多容易出错

### 推荐的替代方案

| 禁用操作 | 替代方案 |
|----------|----------|
| `body.clear()` | 在新文档中操作，或明确告知用户 |
| `insertParagraph(..., "Start")` | 始终使用 `"End"` 顺序添加 |
| `search()` 定位 | 保存引用，使用 `insertParagraph(..., "After")` |
| `insertField(toc)` | 手动创建目录列表，或提示用户使用 Word 内置功能 |

### 示例：创建报告模板（正确的分步方式）

**用户请求**: "创建一个项目报告模板，包含封面、摘要、背景、结论"

**正确做法 - 分 4 步执行**:

**步骤 1/4：创建封面**
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const title = body.insertParagraph("项目报告", "End");
  title.font.size = 28;
  title.font.bold = true;
  title.alignment = Word.Alignment.centered;
  
  const subtitle = body.insertParagraph("[项目名称]", "End");
  subtitle.font.size = 18;
  subtitle.alignment = Word.Alignment.centered;
  
  await context.sync();
  return { success: true, step: "1/4", created: "封面标题" };
});
```

**步骤 2/4：创建摘要章节**
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const heading = body.insertParagraph("1. 摘要", "End");
  heading.style = "Heading 1";
  
  const content = body.insertParagraph("[在此填写摘要内容...]", "End");
  content.font.size = 11;
  
  await context.sync();
  return { success: true, step: "2/4", created: "摘要章节" };
});
```

**步骤 3/4：创建背景章节**
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const heading = body.insertParagraph("2. 背景", "End");
  heading.style = "Heading 1";
  
  const content = body.insertParagraph("[在此填写背景内容...]", "End");
  content.font.size = 11;
  
  await context.sync();
  return { success: true, step: "3/4", created: "背景章节" };
});
```

**步骤 4/4：创建结论章节**
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const heading = body.insertParagraph("3. 结论", "End");
  heading.style = "Heading 1";
  
  const content = body.insertParagraph("[在此填写结论内容...]", "End");
  content.font.size = 11;
  
  await context.sync();
  return { success: true, step: "4/4", created: "结论章节", complete: true };
});
```

### 错误的做法（禁止）

```javascript
// ❌ 错误：一次性创建所有内容（100+ 行代码）
Word.run(async (context) => {
  const body = context.document.body;
  body.clear(); // ❌ 危险操作
  
  // 创建封面...（20 行）
  // 创建摘要...（20 行）
  // 创建背景...（20 行）
  // 创建方法...（20 行）
  // 创建结果...（20 行）
  // 创建结论...（20 行）
  // 添加页眉页脚...（20 行）
  
  await context.sync();
});
```

## 详细模板

更多操作模板请参考 [TOOLS.md](TOOLS.md)。
