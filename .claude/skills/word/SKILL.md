---
name: word-operations
description: Word document operation skill. Used for reading documents, inserting text/images/tables, formatting content, content controls, headers/footers, and document search. Use when users mention Word, documents, paragraphs, tables, or word processing.
---

# Word Operations Skill

## Usage

Manipulate Word by generating **hidden Office.js code** that is automatically executed by the frontend, completely transparent to users.

### Important Rules

- **User Experience First**: Users should only see natural language, not any code
- **Hidden Code Format**: Wrap code with HTML comments: `<!--OFFICE-CODE:word\ncode\n-->`
- **Friendly Feedback**: Inform users of results in natural language after operations complete
- **Complete & Executable**: Generated code must be complete, directly runnable Office.js code

## ⚠️ Tool Selection Priority (Mandatory Rule)

### Prefer MCP Domain Tools

DocuPilot 2.0 provides **Domain-Aggregated MCP Tools** that are faster, safer, and easier to use than the generic execute_code tool.

**Mandatory Rules**:
1. **Use MCP domain tools by default** - Covers 85%+ of common scenarios
2. **Only use execute_code when MCP tools cannot satisfy requirements** - For complex advanced APIs

### Available Word MCP Tools

| Tool | Purpose | Frequency |
|------|---------|-----------|
| `word_paragraph` | Paragraph insert, formatting | ⭐⭐⭐ Most Frequent |
| `word_document` | Document read, search/replace | ⭐⭐ Frequent |
| `word_table` | Table create/edit | ⭐ Medium |
| `execute_code` | Fields, annotations, etc. | Fallback Tool |

### Tool Selection Decision Tree

```
User Request
  |
  ├─ Insert/format paragraphs? → Use word_paragraph
  ├─ Read/search/replace document? → Use word_document
  ├─ Create/edit tables? → Use word_table
  └─ Fields/annotations/headers/footers? → Use execute_code
```

### MCP Tool Invocation Method

```typescript
// ✅ Recommended: Use MCP domain tools
mcp__office__word_paragraph({
  action: "insert",
  text: "Chapter 1: Introduction",
  location: "End",
  format: { style: "Heading 1" }
})

// ❌ Not Recommended: Unless MCP tools cannot meet requirements
mcp__office__execute_code({
  host: "word",
  code: "Word.run(async (context) => { ... })"
})
```

### Example Comparison

**Scenario**: Create a report document

**Using MCP Tools (Recommended)**:
```typescript
// Step 1: Insert heading
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

// Step 2: Insert body text
mcp__office__word_paragraph({
  action: "insert",
  text: "Chapter 1: Overview\n\nThis report analyzes...",
  location: "End",
  format: {
    style: "Normal",
    font: { size: 12 }
  }
})

// Step 3: Insert table
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

**Using execute_code (Only When Necessary)**:
```typescript
// Only when fields (like dynamic date, TOC) are needed
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

### Performance Comparison

| Metric | MCP Tools | execute_code | Improvement |
|--------|-----------|--------------|-------------|
| Response Time | 1.2s | 2.5s | ↓52% |
| Token Cost | ~280 | ~800 | ↓65% |
| Error Rate | <5% | 15% | ↓67% |

### Complete Tool API Reference

For detailed tool parameters and return values, refer to:
- [MCP Tools API Documentation](../../../docs/MCP_TOOLS_API.md)
- [MCP Tools Complete List](../../../docs/MCP_TOOLS_REFERENCE.md)

## Workflow

1. **Understand Requirements**: Analyze user's operation requests
2. **Reference Template**: Check code templates in TOOLS.md
3. **Generate Code**: Create complete Office.js code
4. **Embed Hidden Marker**: Wrap code with `<!--OFFICE-CODE:word ... -->`
5. **Add Friendly Message**: Inform user of operation results

## Supported Features

- **Document Editing**: Insert text, paragraphs, lists, tables, images (Base64).
- **Content Controls**: Create, read, update content controls (for forms/templates).
- **Headers and Footers**: Modify document header and footer content.
- **Formatting**: Set fonts, colors, paragraph spacing, alignment.
- **Search & Replace**: Regular search, wildcard search, batch replace, highlighting.
- **Document Structure**: Operate sections, paragraphs, heading styles.
- **Fields**:
  - Date field: Insert auto-updating current date
  - Table of Contents field (TOC): Auto-generate document TOC
  - Hyperlink field: Create clickable links
  - Page number field: Insert page numbers
  - Addin field: Store custom plugin data
  - Update fields: Refresh field content
- **Footnotes and Endnotes**:
  - Insert footnote references in document
  - Insert endnote references
  - Read and manage footnote/endnote content
- **Style Management**:
  - Apply built-in styles (Heading1-9, Body, Quote, etc.)
  - Use `styleBuiltIn` enum to apply styles
  - Batch apply styles to matching content
  - Get and copy existing paragraph styles
- **Annotations**:
  - Insert grammar and writing suggestion annotations (requires Microsoft 365 subscription)
  - Add colored underline hints to paragraphs
  - Provide popup suggestions and actions
  - Register and respond to annotation events (click, hover, etc.)
  - Delete annotations

## ⚠️ Common Error Handling

### InvalidArgument Error
- **Cause**: Referenced content control, paragraph, or range doesn't exist
- **Solution**: Check if selection or content exists before operating
```javascript
const selection = context.document.getSelection();
selection.load("text");
await context.sync();
if (selection.text.length === 0) {
  console.log("Nothing selected");
  return;
}
```

### InvalidReference Error
- **Cause**: Accessing deleted objects
- **Solution**: Don't continue referencing object after deletion

### Field-Related Errors
- **Issue**: Field code format incorrect
- **Solution**: Use correct field switch format, like date field `\\@ "M/d/yyyy"`
- **Issue**: Field doesn't update
- **Solution**: Call `field.update()` method to manually update field

### Annotation-Related Limitations
- Annotation API requires Microsoft 365 subscription, doesn't support one-time purchase Office versions
- Annotations are not persistently saved to document, need to regenerate after reopening document
- If user accepts suggestion and saves document, changes are retained

### Best Practices
1. Call `load()` and `context.sync()` before reading properties
2. Avoid frequent `context.sync()` calls in loops
3. Use batch operations to improve performance
4. Call `load()` to get results and code after field operations
5. After inserting footnotes/endnotes, can load `reference` property to get reference number

## Examples

### Example 1: Insert Multiple Paragraphs
**User Request**: "Add a chapter titled 'Executive Summary' with introductory text"

**Your Response**:
```
I'll add the chapter heading and introduction for you.

[Step 1: Inserting heading using word_paragraph tool]
[Step 2: Inserting body text using word_paragraph tool]

✅ Successfully added chapter "Executive Summary" with heading and introductory paragraph.
```

**Tools Used**:
1. `mcp__office__word_paragraph({ action: "insert", text: "Executive Summary", location: "End", format: { style: "Heading 1" } })`
2. `mcp__office__word_paragraph({ action: "insert", text: "This report provides...", location: "End", format: { style: "Normal" } })`

### Example 2: Create Data Table
**User Request**: "Create a table showing Q1-Q4 sales figures"

**Your Response**:
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

### Example 3: Insert Poem (Complex Formatting)
**User Request**: "Add Li Bai's 'Quiet Night Thoughts' at end of document"

**Your Response**:
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

**Note**: Complex multi-element formatting with loops is more efficient using execute_code.

### Example 4: Create Content Control Table (Advanced API)
**User Request**: "Create a customer information table with content controls"

**Your Response**:
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

## User File Handling

### Workspace File Access

Users can upload files to current session, these files are saved in workspace directory:

- **User Upload File Path**: `workspace/sessions/{session_id}/uploads/`
- **Generated File Save Path**: `workspace/sessions/{session_id}/outputs/`

### File Operation Flow

1. **Find User-Uploaded Documents**:
   ```typescript
   // Use Glob tool to find Word documents
   // Filenames include timestamp prefix, use wildcards
   const pattern = "workspace/sessions/{session_id}/uploads/*.docx";
   ```

2. **Read Text Files**:
   - For plain text files (TXT, MD), use Read tool to directly read content
   - For Word documents, guide user to open in Word then use Office.js API to operate

3. **Save Processing Results**:
   ```typescript
   // Use Write tool to save processed text
   Write: workspace/sessions/{session_id}/outputs/formatted_text.txt
   ```

### Example Workflow

**User Request**: "Help me format uploaded document"

**Processing Steps**:
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

## Detailed Templates

For more operation templates, please refer to [TOOLS.md](TOOLS.md).
