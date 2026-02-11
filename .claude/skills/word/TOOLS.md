# Word 工具模板库

## ⚠️ 重要：优先使用 MCP 领域工具

**本文件仅包含供参考的低层 Office.js 代码模板。**

**实际开发中请优先使用 MCP 领域工具：**
- `word_paragraph` - 段落插入、格式、删除操作
- `word_document` - 文档读取、搜索、替换操作
- `word_table` - 表格创建、编辑、格式操作

**仅对以下情况使用 execute_code + 本文件模板：**
- 域操作（日期、目录、页码、超链接）
- 脚注和尾注
- 批注和评论
- 高级格式的页眉/页脚
- 内容控件（表单/模板）
- 样式管理（复杂样式操作）
- MCP 工具未覆盖的其他高级 API

**性能对比：**
- MCP 工具：1.2s 响应，~280 tokens，<5% 错误率
- execute_code：2.5s 响应，~800 tokens，15% 错误率

**另见：**
- [MCP 工具 API 文档](../../../docs/MCP_TOOLS_API.md)
- [MCP 工具决策流程](../../../docs/MCP_TOOL_DECISION_FLOW.md)

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

## 文档读取模板

### 读取选中文本
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

### 读取整个文档
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

### 读取文档段落
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

## 文本插入模板

### 在选中处插入文本
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.insertText("Inserted text", "Replace");
  await context.sync();
});
```

### 在文档末尾插入段落
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

### 插入多行文本（推荐）
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  const lines = ["First line content", "Second line content", "Third line content"];
  
  // 循环插入每个段落，确保灵活的格式控制
  for (const line of lines) {
    const p = body.insertParagraph(line, "End");
    // Optional: set paragraph style
    // p.alignment = Word.Alignment.centered;
  }
  
  await context.sync();
});
```

### 在指定位置插入内容
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  // 在开头插入
  body.insertParagraph("Beginning content", "Start");
  
  // 在末尾插入
  body.insertParagraph("End content", "End");
  
  await context.sync();
});
```

## 图片操作模板

### 插入 Base64 图片
```javascript
// Assume base64Image is Base64 string of image (without data:image/... prefix)
Word.run(async (context) => {
  const body = context.document.body;
  
  // 在文档末尾插入图片
  const image = body.insertInlinePictureFromBase64(base64Image, "End");
  
  // 设置图片大小（可选）
  image.width = 400;
  image.height = 300;
  
  await context.sync();
});
```

## 列表操作模板

### 创建列表
```javascript
Word.run(async (context) => {
  const body = context.document.body;
  
  const items = ["Item 1", "Item 2", "Item 3"];
  
  // 插入第一项并开始列表
  const firstPara = body.insertParagraph(items[0], "End");
  firstPara.startNewList();
  
  // 插入后续项
  for (let i = 1; i < items.length; i++) {
    body.insertParagraph(items[i], "End");
  }
  
  await context.sync();
});
```

## 表格操作模板

### 插入表格
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // 创建 3x4 表格
  const table = selection.insertTable(3, 4, "After", [
    ["Header1", "Header2", "Header3", "Header4"],
    ["Data1", "Data2", "Data3", "Data4"],
    ["Data5", "Data6", "Data7", "Data8"]
  ]);
  
  // 设置表格样式
  table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;
  
  await context.sync();
});
```

### 读取表格数据
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

## 内容控件模板

### 创建内容控件
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // 将选中内容包裹在内容控件中
  const contentControl = selection.insertContentControl();
  
  // 设置属性
  contentControl.title = "Customer Name";
  contentControl.tag = "CustomerName";
  contentControl.appearance = Word.ContentControlAppearance.boundingBox;
  contentControl.color = "blue";
  
  await context.sync();
});
```

### 读取/更新内容控件
```javascript
Word.run(async (context) => {
  // 按 Tag 查找控件
  const contentControls = context.document.contentControls.getByTag("CustomerName");
  contentControls.load("items");
  await context.sync();
  
  // 更新所有匹配控件的文本
  for (let cc of contentControls.items) {
    cc.insertText("Contoso Ltd.", "Replace");
  }
  
  await context.sync();
});
```

## 页眉和页脚模板

### 修改页眉
```javascript
Word.run(async (context) => {
  const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
  header.clear();
  
  const paragraph = header.insertParagraph("机密文档 - 仅供内部使用", "Start");
  paragraph.font.color = "red";
  paragraph.alignment = Word.Alignment.centered;
  
  await context.sync();
});
```

## 格式设置模板

### 设置文本格式
```javascript
Word.run(async (context) => {
  const selection = context.document.getSelection();
  
  // 设置字体
  selection.font.name = "Microsoft YaHei";
  selection.font.size = 12;
  selection.font.bold = true;
  selection.font.color = "#333333";
  
  await context.sync();
});
```

### 设置段落格式
```javascript
Word.run(async (context) => {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  for (const paragraph of paragraphs.items) {
    paragraph.lineSpacing = 1.5;  // 1.5 倍行距
    paragraph.spaceAfter = 10;     // 段后间距
    paragraph.alignment = "Justified";  // 两端对齐
  }
  
  await context.sync();
});
```

## 搜索与替换模板

### 简单替换
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

### 高级搜索（通配符）
```javascript
Word.run(async (context) => {
  // 搜索以 'to' 开头、以 'n' 结尾的单词
  const searchResults = context.document.body.search("to*n", {
    matchWildcards: true
  });
  
  searchResults.load("items");
  await context.sync();
  
  // 高亮结果
  for (const result of searchResults.items) {
    result.font.highlightColor = "yellow";
  }
});
```

## 域操作模板

### 插入当前日期域
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 插入日期域，格式「月/日/年 时:分 上午/下午」
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

### 创建目录域
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 插入目录域（TOC）
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

### 插入超链接域
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 插入超链接域
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

### 插入页码域
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 插入页码域
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

### Addin 域（存储插件数据）
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 插入 Addin 域以存储插件数据
  const field = range.insertField(
    Word.InsertLocation.before,
    Word.FieldType.addin
  );
  
  field.load("code,result,data");
  await context.sync();
  
  // 设置域数据
  field.data = JSON.stringify({ customKey: "customValue" });
  await context.sync();
});
```

### 更新域内容
```javascript
Word.run(async (context) => {
  // 获取文档中所有域
  const fields = context.document.body.fields;
  fields.load("items");
  await context.sync();
  
  // 更新所有域
  let updatedCount = 0;
  for (const field of fields.items) {
    try {
      field.load("type");
      await context.sync();
      
      // 更新域内容
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

## 脚注和尾注模板

### 插入脚注
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 在选中位置插入脚注
  const footnote = range.insertFootnote("This is footnote reference content.");
  
  footnote.load("reference");
  await context.sync();
  
  console.log("Footnote reference number:", footnote.reference);
});
```

### 插入尾注
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 在选中位置插入尾注
  const endnote = range.insertEndnote("This is endnote reference content.");
  
  endnote.load("reference");
  await context.sync();
  
  console.log("Endnote reference number:", endnote.reference);
});
```

### 读取脚注内容
```javascript
Word.run(async (context) => {
  // 搜索脚注引用标记
  const searchResults = context.document.body.search("^f", {
    matchWildcards: true
  });
  
  searchResults.load("items");
  await context.sync();
  
  console.log("Found", searchResults.items.length, "footnotes");
});
```

## 样式管理模板

### 应用标题 1 样式
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection().getRange();
  
  // 应用内置标题 1 样式
  range.style = "Heading1";
  
  await context.sync();
});
```

### 应用标题样式（使用枚举）
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection();
  range.load("text");
  await context.sync();
  
  // 使用样式的字符串形式（兼容性更好）
  // 选项："Heading 1"、"Heading 2"、"Heading 3" 等
  range.style = "Heading 2";
  
  await context.sync();
  
  return {
    success: true,
    appliedStyle: "Heading 2",
    text: range.text
  };
});
```

### 应用引用样式
```javascript
Word.run(async (context) => {
  const range = context.document.getSelection();
  range.load("text");
  await context.sync();
  
  // 应用引用块样式（使用字符串形式）
  range.style = "Quote";
  
  await context.sync();
  
  return {
    success: true,
    appliedStyle: "Quote",
    text: range.text
  };
});
```

### 批量应用样式
```javascript
Word.run(async (context) => {
  // 查找包含指定文本的所有段落
  const searchResults = context.document.body.search("Important", {
    matchCase: false
  });
  
  searchResults.load("items");
  await context.sync();
  
  let appliedCount = 0;
  
  // 对所有匹配段落应用强调样式
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

### 获取并应用现有样式
```javascript
Word.run(async (context) => {
  // 获取第一段的样式
  const firstParagraph = context.document.body.paragraphs.getFirst();
  firstParagraph.load("style");
  await context.sync();
  
  const styleName = firstParagraph.style;
  
  // 将该样式应用于选中区域
  const range = context.document.getSelection().getRange();
  range.style = styleName;
  
  await context.sync();
});
```

## 批注模板

### 插入语法批注
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // 创建批注（需要 WordApi 1.7+）
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

### 读取段落批注
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // 获取段落的所有批注
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

### 注册批注事件
```javascript
Word.run(async (context) => {
  // 注册批注点击事件
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

### 删除批注
```javascript
Word.run(async (context) => {
  const paragraph = context.document.getSelection().paragraphs.getFirst();
  
  // 获取并删除段落的所有批注
  const annotations = paragraph.getAnnotations();
  annotations.load("items");
  await context.sync();
  
  for (const annotation of annotations.items) {
    annotation.delete();
  }
  
  await context.sync();
});
```

## 文档处理 Python 模板

### 文本摘要
```python
def summarize_text(text, max_length=200):
    """使用 AI 摘要文本"""
    # In actual implementation, call Claude API
    prompt = f"""Please summarize the following text, keeping it within {max_length} words:

{text}

Summary:"""
    # Call AI to generate summary
    return summary
```

### 文本改写
```python
def rewrite_text(text, style="formal"):
    """改写文本风格"""
    styles = {
        "formal": "formal, professional tone",
        "casual": "relaxed, friendly tone",
        "concise": "concise, refined expression",
        "detailed": "detailed, comprehensive description"
    }
    
    prompt = f"""Please rewrite the following text in a {styles[style]}:

Original: {text}

Rewritten:"""
    # 调用 AI 改写
    return rewritten_text
```
