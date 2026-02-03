# Word Tool Template Library

## ⚠️ Important: Prefer MCP Domain Tools

**This file contains low-level Office.js code templates for reference only.**

**In actual development, prefer MCP domain tools:**
- `word_paragraph` - Paragraph insert, format, delete operations
- `word_document` - Document read, search, replace operations
- `word_table` - Table create, edit, format operations

**Only use execute_code + templates in this file for:**
- Field operations (Date, TOC, Page numbers, Hyperlinks)
- Footnotes and Endnotes
- Annotations and Comments
- Headers/Footers with advanced formatting
- Content Controls (forms/templates)
- Style Management (complex style operations)
- Other advanced APIs not covered by MCP tools

**Performance Comparison:**
- MCP Tools: 1.2s response, ~280 tokens, <5% error rate
- execute_code: 2.5s response, ~800 tokens, 15% error rate

**See Also:**
- [MCP Tools API Documentation](../../../docs/MCP_TOOLS_API.md)
- [MCP Tools Decision Flow](../../../docs/MCP_TOOL_DECISION_FLOW.md)

---

## 🚨 API 稳定性指南（必读）

### 禁用 API 清单

以下 API 在实际使用中容易导致静默失败或不可预期行为，**应避免使用**：

| API | 风险等级 | 问题描述 | 替代方案 |
|-----|---------|---------|---------|
| `body.clear()` | 🔴 高危 | 清空整个文档，后续操作可能失败 | 在空白文档开始，或明确告知用户 |
| `insertParagraph(..., "Start")` | 🔴 高危 | 在开头插入会打乱已有结构 | 始终使用 `"End"` 顺序添加 |
| `insertField(toc)` | 🟡 中危 | 目录字段不稳定，参数复杂 | 手动创建目录列表，或提示用户用 Word 内置功能 |
| `insertField(page)` 在页脚 | 🟡 中危 | 页码字段在某些环境不工作 | 使用纯文本占位符 |
| `search().insertParagraph("After")` | 🟡 中危 | 依赖搜索结果定位，前置步骤失败则无法定位 | 保存段落引用，使用 `paragraph.insertParagraph("After")` |
| `shading.backgroundPatternColor` | 🟡 中危 | 某些 Word 版本不支持 | 使用 `font.highlightColor` 代替 |

### 可靠代码模板（推荐）

#### 模板 1：单个章节创建（最可靠）

```javascript
// ✅ 推荐：每次只创建一个章节
Word.run(async (context) => {
  const body = context.document.body;
  
  // 创建章节标题
  const heading = body.insertParagraph("章节标题", "End");
  heading.style = "Heading 1";
  heading.spaceAfter = 12;
  
  // 创建章节内容（最多 3-4 段）
  const content = body.insertParagraph("章节内容...", "End");
  content.font.size = 11;
  content.lineSpacing = 1.5;
  
  await context.sync();
  
  // 必须返回验证信息
  return {
    success: true,
    created: "1 个标题 + 1 个内容段落",
    sectionName: "章节标题"
  };
});
```

#### 模板 2：带验证的表格创建

```javascript
// ✅ 推荐：创建表格并验证
Word.run(async (context) => {
  const body = context.document.body;
  
  // 添加表格标题
  const caption = body.insertParagraph("表 1：数据汇总", "End");
  caption.font.bold = true;
  caption.spaceAfter = 6;
  
  // 创建简单表格（建议不超过 5x5）
  const table = body.insertTable(3, 3, "End", [
    ["列1", "列2", "列3"],
    ["数据1", "数据2", "数据3"],
    ["数据4", "数据5", "数据6"]
  ]);
  
  // 设置表格样式（使用可靠的内置样式）
  table.styleBuiltIn = Word.BuiltInStyleName.gridTable4Accent1;
  
  await context.sync();
  
  return {
    success: true,
    created: "3x3 表格",
    tableCaption: "表 1：数据汇总"
  };
});
```

#### 模板 3：安全的页眉设置

```javascript
// ✅ 推荐：安全的页眉设置方式
Word.run(async (context) => {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  
  if (sections.items.length > 0) {
    const header = sections.items[0].getHeader(Word.HeaderFooterType.primary);
    
    // 不要用 header.clear()，直接插入内容
    const headerPara = header.insertParagraph("文档标题 - 页眉", "End");
    headerPara.font.size = 9;
    headerPara.font.color = "#666666";
    headerPara.alignment = Word.Alignment.centered;
    
    await context.sync();
  }
  
  return {
    success: true,
    created: "页眉"
  };
});
```

#### 模板 4：分步创建报告的标准流程

```javascript
// 步骤 1：封面（单独执行）
Word.run(async (context) => {
  const body = context.document.body;
  
  const title = body.insertParagraph("报告标题", "End");
  title.font.size = 28;
  title.font.bold = true;
  title.alignment = Word.Alignment.centered;
  title.spaceAfter = 20;
  
  const subtitle = body.insertParagraph("[副标题]", "End");
  subtitle.font.size = 16;
  subtitle.alignment = Word.Alignment.centered;
  subtitle.spaceAfter = 40;
  
  const author = body.insertParagraph("作者：[姓名]", "End");
  author.alignment = Word.Alignment.centered;
  
  const date = body.insertParagraph("日期：[YYYY-MM-DD]", "End");
  date.alignment = Word.Alignment.centered;
  
  await context.sync();
  return { success: true, step: "1/N", created: "封面" };
});

// 步骤 2-N：各章节（每个章节单独执行）
// ... 参考模板 1
```

### 验证返回值规范

每次代码执行**必须**返回以下格式的验证信息：

```typescript
interface ExecutionResult {
  success: boolean;           // 是否成功
  step?: string;              // 当前步骤，如 "1/4"
  created: string;            // 创建了什么，如 "封面标题"
  paragraphCount?: number;    // 创建的段落数
  tableCount?: number;        // 创建的表格数
  preview?: string;           // 内容预览（前 50 字符）
  complete?: boolean;         // 是否是最后一步
}
```

### 代码复杂度自检清单

在提交代码前，检查以下项目：

- [ ] 代码行数 ≤ 30 行
- [ ] `insert*` 操作 ≤ 5 次
- [ ] 没有使用 `body.clear()`
- [ ] 没有使用 `insertParagraph(..., "Start")`
- [ ] 没有使用复杂的 `search()` 定位
- [ ] 包含 `return { success: true, ... }` 验证返回
- [ ] 只处理一个逻辑单元（如一个章节）

---

## Document Reading Templates

### Read Selected Text
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text");
  await context.sync();
  
  return {
    text: selection.text
  };
});
```

### Read Entire Document
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  
  return {
    text: body.text
  };
});
```

### Read Document Paragraphs
```javascript
Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  const texts = paragraphs.items.map(p => {
    p.load("text");
    return p;
  });
  await context.sync();
  
  return texts.map(p => p.text);
});
```

## Text Insertion Templates

### Insert Text at Selection
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.insertText("Inserted text", "Replace");
  await context.sync();
});
```

### Insert Paragraph at End of Document
```javascript
Word.run(async (context) => {
  const paragraph = context.document.body.insertParagraph(
    "This is new paragraph content",
    "End"
  );
  paragraph.load("text");
  await context.sync();
});
```

### Insert Multiple Lines of Text (Recommended)
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  const lines = ["First line content", "Second line content", "Third line content"];
  
  // Loop to insert each paragraph, ensuring flexible format control
  for (const line of lines) {
    const p = body.insertParagraph(line, "End");
    // Optional: set paragraph style
    // p.alignment = Word.Alignment.centered;
  }
  
  await context.sync();
});
```

### Insert Content at Specific Position
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  // Insert at beginning
  body.insertParagraph("Beginning content", "Start");
  
  // Insert at end
  body.insertParagraph("End content", "End");
  
  await context.sync();
});
```

## Image Operation Templates

### Insert Base64 Image
```javascript
// Assume base64Image is Base64 string of image (without data:image/... prefix)
Word.run(async (context) => {
  const body = context.document.body;
  
  // Insert image at end of document
  const image = body.insertInlinePictureFromBase64(base64Image, "End");
  
  // Set image size (optional)
  image.width = 400;
  image.height = 300;
  
  await context.sync();
});
```

## List Operation Templates

### Create List
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const items = ["Item 1", "Item 2", "Item 3"];
  
  // Insert first item and start list
  const firstPara = body.insertParagraph(items[0], "End");
  firstPara.startNewList();
  
  // Insert subsequent items
  for (let i = 1; i < items.length; i++) {
    body.insertParagraph(items[i], "End");
  }
  
  await context.sync();
});
```

## Table Operation Templates

### Insert Table
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // Create 3x4 table
  const table = selection.insertTable(3, 4, "After", [
    ["Header1", "Header2", "Header3", "Header4"],
    ["Data1", "Data2", "Data3", "Data4"],
    ["Data5", "Data6", "Data7", "Data8"]
  ]);
  
  // Set table style
  table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;
  
  await context.sync();
});
```

### Read Table Data
```javascript
Word.run(async (context) => {
  const tables = context.document.body.tables;
  tables.load("items");
  await context.sync();
  
  if (tables.items.length > 0) {
    const table = tables.items[0];
    const rows = table.rows;
    rows.load("items");
    await context.sync();
    
    const data = [];
    for (const row of rows.items) {
      const cells = row.cells;
      cells.load("items");
      await context.sync();
      
      const rowData = [];
      for (const cell of cells.items) {
        cell.body.load("text");
        await context.sync();
        rowData.push(cell.body.text);
      }
      data.push(rowData);
    }
    
    return data;
  }
});
```

## Content Control Templates

### Create Content Control
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // Wrap selected content in content control
  const contentControl = selection.insertContentControl();
  
  // Set properties
  contentControl.title = "Customer Name";
  contentControl.tag = "CustomerName";
  contentControl.appearance = Word.ContentControlAppearance.boundingBox;
  contentControl.color = "blue";
  
  await context.sync();
});
```

### Read/Update Content Control
```javascript
Word.run(async (context) => {
  // Find control by Tag
  const contentControls = context.document.contentControls.getByTag("CustomerName");
  contentControls.load("items");
  await context.sync();
  
  // Update text of all matching controls
  for (let cc of contentControls.items) {
    cc.insertText("Contoso Ltd.", "Replace");
  }
  
  await context.sync();
});
```

## Header and Footer Templates

### Modify Header
```javascript
Word.run(async (context) => {
  const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
  header.clear();
  
  const paragraph = header.insertParagraph("Confidential Document - Internal Use Only", "Start");
  paragraph.font.color = "red";
  paragraph.alignment = Word.Alignment.centered;
  
  await context.sync();
});
```

## Formatting Templates

### Set Text Format
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // Set font
  selection.font.name = "Microsoft YaHei";
  selection.font.size = 12;
  selection.font.bold = true;
  selection.font.color = "#333333";
  
  await context.sync();
});
```

### Set Paragraph Format
```javascript
Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  for (const paragraph of paragraphs.items) {
    paragraph.lineSpacing = 1.5;  // 1.5x line spacing
    paragraph.spaceAfter = 10;     // Space after paragraph
    paragraph.alignment = "Justified";  // Justified alignment
  }
  
  await context.sync();
});
```

## Search and Replace Templates

### Simple Replace
```javascript
Word.run(async (context) => {
  const searchResults = context.document.body.search("old text", {
    matchCase: false,
    matchWholeWord: false
  });
  
  searchResults.load("items");
  await context.sync();
  
  for (const result of searchResults.items) {
    result.insertText("new text", "Replace");
  }
  
  await context.sync();
  
  return { replacedCount: searchResults.items.length };
});
```

### Advanced Search (Wildcards)
```javascript
Word.run(async (context) => {
  // Search for words starting with 'to' and ending with 'n'
  const searchResults = context.document.body.search("to*n", {
    matchWildcards: true
  });
  
  searchResults.load("items");
  await context.sync();
  
  // Highlight results
  for (const result of searchResults.items) {
    result.font.highlightColor = "yellow";
  }
});
```

## Field Operation Templates

### Insert Current Date Field
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert date field, format "Month/Day/Year Hour:Minute AM/PM"
  const field = range.insertField(
    Word.InsertLocation.end,
    Word.FieldType.date,
    '\\@ "M/d/yyyy h:mm am/pm"',
    true
  );
  
  field.load("result,code");
  await context.sync();
  
  console.log("Date field code:", field.code);
  console.log("Date field result:", field.result);
});
```

### Create Table of Contents Field
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert Table of Contents field (TOC)
  const field = range.insertField(
    Word.InsertLocation.start,
    Word.FieldType.toc,
    '\\o "1-3" \\h \\z \\u',
    true
  );
  
  field.load("result");
  await context.sync();
});
```

### Insert Hyperlink Field
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert hyperlink field
  const field = range.insertField(
    Word.InsertLocation.end,
    Word.FieldType.hyperlink,
    '"https://www.microsoft.com" \\o "Visit Microsoft"',
    true
  );
  
  field.load("result,code");
  await context.sync();
});
```

### Insert Page Number Field
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert page number field
  const field = range.insertField(
    Word.InsertLocation.end,
    Word.FieldType.page,
    "",
    true
  );
  
  field.load("result");
  await context.sync();
});
```

### Addin Field (Store Plugin Data)
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert Addin field to store plugin data
  const field = range.insertField(
    Word.InsertLocation.before,
    Word.FieldType.addin
  );
  
  field.load("code,result,data");
  await context.sync();
  
  // Set field data
  field.data = JSON.stringify({ customKey: "customValue" });
  await context.sync();
});
```

### Update Field Content
```javascript
Word.run(async (context) => {
  // Get all fields in document
  const fields = context.document.body.fields;
  fields.load("items");
  await context.sync();
  
  // Update all fields
  let updatedCount = 0;
  for (const field of fields.items) {
    try {
      field.load("type");
      await context.sync();
      
      // Update field content
      field.updateResult();
      updatedCount++;
    } catch (e) {
      console.warn("Cannot update field:", e);
    }
  }
  
  await context.sync();
  
  return {
    success: true,
    updatedCount: updatedCount,
    message: `Successfully updated ${updatedCount} fields`
  };
});
```

## Footnote and Endnote Templates

### Insert Footnote
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert footnote at selected position
  const footnote = range.insertFootnote("This is footnote reference content.");
  
  footnote.load("reference");
  await context.sync();
  
  console.log("Footnote reference number:", footnote.reference);
});
```

### Insert Endnote
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Insert endnote at selected position
  const endnote = range.insertEndnote("This is endnote reference content.");
  
  endnote.load("reference");
  await context.sync();
  
  console.log("Endnote reference number:", endnote.reference);
});
```

### Read Footnote Content
```javascript
Word.run(async (context) => {
  // Search for footnote reference markers
  const searchResults = context.document.body.search("^f", {
    matchWildcards: true
  });
  
  searchResults.load("items");
  await context.sync();
  
  console.log("Found", searchResults.items.length, "footnotes");
});
```

## Style Management Templates

### Apply Heading 1 Style
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // Apply built-in Heading 1 style
  range.style = "Heading1";
  
  await context.sync();
});
```

### Apply Heading Style (Using Enum)
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection();
  range.load("text");
  await context.sync();
  
  // Use string form of style name (more compatible)
  // Options: "Heading 1", "Heading 2", "Heading 3", etc.
  range.style = "Heading 2";
  
  await context.sync();
  
  return {
    success: true,
    appliedStyle: "Heading 2",
    text: range.text
  };
});
```

### Apply Quote Style
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection();
  range.load("text");
  await context.sync();
  
  // Apply quote block style (using string form)
  range.style = "Quote";
  
  await context.sync();
  
  return {
    success: true,
    appliedStyle: "Quote",
    text: range.text
  };
});
```

### Batch Apply Styles
```javascript
Word.run(async (context) => {
  // Find all paragraphs containing specific text
  const searchResults = context.document.body.search("Important", {
    matchCase: false
  });
  
  searchResults.load("items");
  await context.sync();
  
  let appliedCount = 0;
  
  // Apply emphasis style to all matching paragraphs
  for (const result of searchResults.items) {
    const paragraph = result.parentOrNullObject.paragraph;
    paragraph.load("isNullObject");
    await context.sync();
    
    if (!paragraph.isNullObject) {
      paragraph.style = "Emphasis";  // Use string form
      appliedCount++;
    }
  }
  
  await context.sync();
  
  return {
    success: true,
    appliedCount: appliedCount,
    message: `Successfully applied emphasis style to ${appliedCount} paragraphs`
  };
});
```

### Get and Apply Existing Style
```javascript
Word.run(async (context) => {
  // Get first paragraph's style
  const firstParagraph = context.document.body.paragraphs.getFirst();
  firstParagraph.load("style");
  await context.sync();
  
  const styleName = firstParagraph.style;
  
  // Apply that style to selected range
  const range = context.document.getSelection().getRange();
  range.style = styleName;
  
  await context.sync();
});
```

## Annotation Templates

### Insert Grammar Annotation
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // Create annotation (requires WordApi 1.7+)
  const annotationSet = paragraph.insertAnnotations([{
    critiqueAnnotation: {
      critique: {
        colorScheme: Word.CritiqueColorScheme.red,
        start: 0,
        length: 10
      },
      popupOptions: {
        title: "Grammar Suggestion",
        suggestions: ["Suggestion 1", "Suggestion 2"],
        subtitle: "Possible grammar issue"
      }
    }
  }]);
  
  await context.sync();
});
```

### Read Paragraph Annotations
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // Get all annotations of paragraph
  const annotations = paragraph.getAnnotations();
  annotations.load("items");
  await context.sync();
  
  console.log("Found", annotations.items.length, "annotations");
  
  for (const annotation of annotations.items) {
    annotation.load("critiqueAnnotation");
    await context.sync();
    console.log("Annotation:", annotation.critiqueAnnotation);
  }
});
```

### Register Annotation Event
```javascript
Word.run(async (context) => {
  // Register annotation click event
  context.document.onAnnotationClicked.add(async (args) => {
    await Word.run(async (context) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");
      await context.sync();
      
      console.log("Clicked annotation:", annotation.critiqueAnnotation.critique);
    });
  });
  
  await context.sync();
});
```

### Delete Annotation
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // Get and delete all annotations of paragraph
  const annotations = paragraph.getAnnotations();
  annotations.load("items");
  await context.sync();
  
  for (const annotation of annotations.items) {
    annotation.delete();
  }
  
  await context.sync();
});
```

## Document Processing Python Templates

### Text Summarization
```python
def summarize_text(text, max_length=200):
    """Summarize text using AI"""
    # In actual implementation, call Claude API
    prompt = f"""Please summarize the following text, keeping it within {max_length} words:

{text}

Summary:"""
    # Call AI to generate summary
    return summary
```

### Text Rewriting
```python
def rewrite_text(text, style="formal"):
    """Rewrite text style"""
    styles = {
        "formal": "formal, professional tone",
        "casual": "relaxed, friendly tone",
        "concise": "concise, refined expression",
        "detailed": "detailed, comprehensive description"
    }
    
    prompt = f"""Please rewrite the following text in a {styles[style]}:

Original: {text}

Rewritten:"""
    # Call AI to rewrite
    return rewritten_text
```
