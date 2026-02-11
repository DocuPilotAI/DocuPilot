# PowerPoint 工具模板库

## ⚠️ 重要：优先使用 MCP 领域工具

**本文件仅包含供参考的低层 Office.js 代码模板。**

**实际开发中请优先使用 MCP 领域工具：**
- `ppt_shape` - 文本框、图片、几何形状操作
- `ppt_slide` - 幻灯片管理（添加、删除、复制、移动）
- `ppt_table` - 表格创建、编辑、格式操作

**仅对以下情况使用 execute_code + 本文件模板：**
- 动画和切换效果
- 幻灯片母版和主题操作
- 形状分组操作
- 高级文本格式（文本框边距、垂直对齐）
- 需要特定模板的幻灯片版式
- MCP 工具未覆盖的其他高级 API

**性能对比：**
- MCP 工具：1.2s 响应，~280 tokens，<5% 错误率
- execute_code：2.5s 响应，~800 tokens，15% 错误率

**另见：**
- [MCP 工具 API 文档](../../../docs/MCP_TOOLS_API.md)
- [MCP 工具决策流程](../../../docs/MCP_TOOL_DECISION_FLOW.md)

---

## 幻灯片操作模板

### 获取幻灯片列表
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  return {
    slideCount: slides.items.length,
    slides: slides.items.map((slide, index) => ({
      index: index + 1,
      id: slide.id
    }))
  };
});
```

### 添加空白幻灯片
```javascript
PowerPoint.run(async (context) => {
  context.presentation.slides.add();
  await context.sync();
});
```

### 添加幻灯片并添加内容（重要！）
**关键**：添加新幻灯片并立即向其添加形状时，必须在 add() 后 sync()，然后获取幻灯片引用：

```javascript
PowerPoint.run(async (context) => {
  // 步骤 1：添加新幻灯片
  context.presentation.slides.add();
  await context.sync(); // 关键：访问新幻灯片前必须 sync
  
  // 步骤 2：获取新添加幻灯片的引用
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  const newSlide = slides.items[slides.items.length - 1];
  
  // 步骤 3：现在可以向新幻灯片添加形状
  const textBox = newSlide.shapes.addTextBox("Hello World");
  textBox.left = 100;
  textBox.top = 100;
  textBox.width = 300;
  textBox.height = 50;
  
  await context.sync();
});
```

**替代方式 - 使用 getSelectedSlides()：**
```javascript
PowerPoint.run(async (context) => {
  // 若操作当前选中的幻灯片
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 可立即添加形状，无需额外 sync
  const textBox = slide.shapes.addTextBox("Hello World");
  textBox.left = 100;
  textBox.top = 100;
  
  await context.sync();
});
```

### 删除幻灯片
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  // 删除第一张幻灯片
  if (slides.items.length > 0) {
    slides.items[0].delete();
  }
  
  await context.sync();
});
```

### 从其他演示文稿插入幻灯片
```javascript
// 假设 chosenFileBase64 是包含 .pptx 文件内容的 Base64 字符串
PowerPoint.run(async (context) => {
  context.presentation.insertSlidesFromBase64(chosenFileBase64);
  await context.sync();
});
```

## 表格操作模板

### 创建并填充表格
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 创建 3x4 表格
  const table = slide.shapes.addTable(3, 4, {
    left: 100,
    top: 100,
    width: 400,
    height: 150,
    values: [
      ["Header1", "Header2", "Header3", "Header4"],
      ["Data1", "Data2", "Data3", "Data4"],
      ["Data5", "Data6", "Data7", "Data8"]
    ]
  });
  
  await context.sync();
});
```

### 格式化表格单元格
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 创建表格并应用样式
  const table = slide.shapes.addTable(3, 3, {
    left: 100,
    top: 100,
    values: [
      ["A", "B", "C"],
      ["D", "E", "F"],
      ["G", "H", "I"]
    ],
    uniformCellProperties: {
      fill: { color: "darkblue" },
      font: { color: "white", bold: true },
      horizontalAlignment: PowerPoint.ParagraphHorizontalAlignment.center
    }
  });
  
  await context.sync();
});
```

## 形状操作模板

### 添加矩形
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const rect = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect.left = 100;
  rect.top = 100;
  rect.width = 200;
  rect.height = 100;
  rect.fill.setSolidColor("blue");
  
  await context.sync();
});
```

### 添加文本框
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const textBox = slide.shapes.addTextBox("This is text content");
  textBox.left = 100;
  textBox.top = 100;
  textBox.width = 300;
  textBox.height = 50;
  
  // 设置文本样式
  textBox.textFrame.textRange.font.size = 24;
  textBox.textFrame.textRange.font.color = "purple";
  
  await context.sync();
});
```

### 添加线条
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: 50,
    top: 50,
    width: 200, // 起点到终点的水平距离
    height: 200 // 起点到终点的垂直距离
  });
  
  await context.sync();
});
```

### 组合形状
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  
  // 假设要组合前两个形状
  slide.load("shapes/items");
  await context.sync();
  
  if (shapes.items.length >= 2) {
    const shapesToGroup = [shapes.items[0], shapes.items[1]];
    const group = shapes.addGroup(shapesToGroup);
  }
  
  await context.sync();
});
```

## 图片操作模板

### 添加图片
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // base64Image 是图片的 Base64 字符串（不含 data:image/... 前缀）
  const image = slide.shapes.addImage(base64Image);
  image.left = 100;
  image.top = 100;
  image.width = 400;
  image.height = 300;
  
  await context.sync();
});
```

## 标签（元数据）模板

### 向幻灯片添加标签
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  slide.tags.add("CATEGORY", "Confidential");
  await context.sync();
});
```

### 按标签处理幻灯片
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("tags/key, tags/value");
  await context.sync();
  
  // 遍历并处理
  for (let slide of slides.items) {
    // 检查标签逻辑
  }
});
```

## 文本格式增强模板

### 设置文本粗体和斜体
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // 设置文本范围字体属性
  const textRange = shape.textFrame.textRange;
  textRange.font.bold = true;
  textRange.font.italic = true;
  textRange.font.underline = PowerPoint.FontUnderlineStyle.single;
  
  await context.sync();
});
```

### 设置字体名称和大小
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  const textRange = shape.textFrame.textRange;
  textRange.font.name = "Microsoft YaHei";
  textRange.font.size = 24;
  textRange.font.color = "#4472C4";
  
  await context.sync();
});
```

### 文本垂直居中
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // 设置文本框垂直对齐
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  
  await context.sync();
});
```

### 设置文本框边距
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // 设置文本框边距（单位：磅）
  shape.textFrame.marginLeft = 10;
  shape.textFrame.marginRight = 10;
  shape.textFrame.marginTop = 5;
  shape.textFrame.marginBottom = 5;
  
  await context.sync();
});
```

## 幻灯片版式和母版模板

### 获取幻灯片版式信息
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 加载版式和母版信息
  slide.load("slideMaster,layout");
  await context.sync();
  
  const master = slide.slideMaster;
  const layout = slide.layout;
  
  master.load("id,name");
  layout.load("id,name");
  await context.sync();
  
  console.log("Master:", master.name);
  console.log("Layout:", layout.name);
});
```

### 使用指定版式创建幻灯片
```javascript
PowerPoint.run(async (context) => {
  // 获取第一张幻灯片的版式
  const firstSlide = context.presentation.slides.getItemAt(0);
  firstSlide.load("layout");
  await context.sync();
  
  const layout = firstSlide.layout;
  layout.load("id");
  await context.sync();
  
  // 使用相同版式创建新幻灯片
  const newSlide = context.presentation.slides.add({
    layoutId: layout.id
  });
  
  await context.sync();
});
```

### 获取所有可用版式
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (slides.items.length > 0) {
    const firstSlide = slides.items[0];
    firstSlide.load("slideMaster");
    await context.sync();
    
    const master = firstSlide.slideMaster;
    master.load("layouts");
    await context.sync();
    
    const layouts = master.layouts;
    layouts.load("items");
    await context.sync();
    
    console.log("可用版式：");
    for (const layout of layouts.items) {
      layout.load("id,name");
      await context.sync();
      console.log(`- ${layout.name} (ID: ${layout.id})`);
    }
  }
});
```

## 主题系统模板

### 用主题颜色填充形状
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // 使用主题颜色（accent1 为主题强调色 1）
  shape.fill.setSolidColor("accent1");
  
  await context.sync();
});
```

### 使用多种主题颜色
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // 使用不同主题颜色创建多个形状
  const rect1 = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect1.left = 50;
  rect1.top = 50;
  rect1.width = 100;
  rect1.height = 100;
  rect1.fill.setSolidColor("accent1"); // 主题强调色 1
  
  const rect2 = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect2.left = 170;
  rect2.top = 50;
  rect2.width = 100;
  rect2.height = 100;
  rect2.fill.setSolidColor("accent2"); // 主题强调色 2
  
  await context.sync();
});
```

## 表格数据操作模板

### 读取表格数据
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  
  // 查找第一个表格
  let table = null;
  for (const shape of shapes.items) {
    if (shape.type === PowerPoint.ShapeType.table) {
      table = shape.table;
      break;
    }
  }
  
  if (table) {
    table.load("rows,columns");
    await context.sync();
    
    // 读取所有单元格数据
    for (let i = 0; i < table.rows.count; i++) {
      for (let j = 0; j < table.columns.count; j++) {
        const cell = table.getCell(i, j);
        cell.load("text");
        await context.sync();
        console.log(`[${i},${j}]: ${cell.text}`);
      }
    }
  }
});
```

### 更新表格单元格
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  
  let table = null;
  for (const shape of shapes.items) {
    if (shape.type === PowerPoint.ShapeType.table) {
      table = shape.table;
      break;
    }
  }
  
  if (table) {
    // 更新单元格内容
    const cell = table.getCell(0, 0);
    cell.text = "Updated content";
    
    await context.sync();
  }
});
```

### 设置行高和列宽
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  
  let table = null;
  for (const shape of shapes.items) {
    if (shape.type === PowerPoint.ShapeType.table) {
      table = shape.table;
      break;
    }
  }
  
  if (table) {
    // 设置行高和列宽
    const firstRow = table.rows.getItemAt(0);
    firstRow.height = 50;
    
    const firstColumn = table.columns.getItemAt(0);
    firstColumn.width = 150;
    
    await context.sync();
  }
});
```

### 创建含合并单元格的表格
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const tableValues = [
    ["Title", "", "", ""],  // 第一行将被合并
    ["Item A", "100", "200", "300"],
    ["Item B", "150", "250", "350"]
  ];
  
  const options = {
    values: tableValues,
    left: 50,
    top: 50,
    mergedAreas: [{
      firstRow: 0,
      firstColumn: 0,
      rowSpan: 1,
      columnSpan: 4
    }]
  };
  
  const table = slide.shapes.addTable(3, 4, options);
  await context.sync();
});
```

## 演示文稿生成 Python 模板

### 生成演示大纲
```python
def generate_outline(topic, slide_count=5):
    """根据主题生成演示大纲"""
    prompt = f"""Please generate a {slide_count}-page presentation outline for the following topic:

Topic: {topic}

Please output in the following format:
1. Title Slide: [Title]
2. Content Slide 1: [Title] - [Point1, Point2, Point3]
3. Content Slide 2: [Title] - [Point1, Point2, Point3]
...
{slide_count}. Summary Slide: [Key Conclusions]

Outline:"""
    
    # 调用 AI 生成大纲
    return outline
```

### 生成幻灯片内容
```python
def generate_slide_content(title, points):
    """生成幻灯片的详细内容"""
    content = {
        "title": title,
        "bullet_points": points,
        "speaker_notes": generate_speaker_notes(title, points)
    }
    return content

def generate_speaker_notes(title, points):
    """生成演讲备注"""
    prompt = f"""Please generate concise speaker notes for the following slide:

Title: {title}
Points: {', '.join(points)}

Speaker Notes (50-100 words):"""
    
    # 调用 AI 生成备注
    return notes
```

### 配色方案建议
```python
COLOR_SCHEMES = {
    "professional": {
        "primary": "#2E5090",
        "secondary": "#4472C4",
        "accent": "#ED7D31",
        "text": "#333333",
        "background": "#FFFFFF"
    },
    "modern": {
        "primary": "#1A1A2E",
        "secondary": "#16213E",
        "accent": "#E94560",
        "text": "#EAEAEA",
        "background": "#0F3460"
    },
    "nature": {
        "primary": "#2D5A27",
        "secondary": "#4A7C45",
        "accent": "#F4A259",
        "text": "#333333",
        "background": "#F5F5F5"
    }
}

def get_color_scheme(style="professional"):
    """获取配色方案"""
    return COLOR_SCHEMES.get(style, COLOR_SCHEMES["professional"])
```
