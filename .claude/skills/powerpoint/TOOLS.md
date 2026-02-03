# PowerPoint Tool Template Library

## Slide Operation Templates

### Get Slide List
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

### Add Blank Slide
```javascript
PowerPoint.run(async (context) => {
  context.presentation.slides.add();
  await context.sync();
});
```

### Add Slide and Add Content (IMPORTANT!)
**Critical**: When adding a new slide and immediately adding shapes to it, you MUST sync() after add() and then get the slide reference:

```javascript
PowerPoint.run(async (context) => {
  // Step 1: Add new slide
  context.presentation.slides.add();
  await context.sync(); // CRITICAL: Must sync before accessing the new slide
  
  // Step 2: Get reference to the newly added slide
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  const newSlide = slides.items[slides.items.length - 1];
  
  // Step 3: Now you can add shapes to the new slide
  const textBox = newSlide.shapes.addTextBox("Hello World");
  textBox.left = 100;
  textBox.top = 100;
  textBox.width = 300;
  textBox.height = 50;
  
  await context.sync();
});
```

**Alternative approach - Use getSelectedSlides():**
```javascript
PowerPoint.run(async (context) => {
  // If working with currently selected slide
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // Can immediately add shapes without extra sync
  const textBox = slide.shapes.addTextBox("Hello World");
  textBox.left = 100;
  textBox.top = 100;
  
  await context.sync();
});
```

### Delete Slide
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  // Delete first slide
  if (slides.items.length > 0) {
    slides.items[0].delete();
  }
  
  await context.sync();
});
```

### Insert Slides from Another Presentation
```javascript
// Assume chosenFileBase64 is Base64 string containing .pptx file content
PowerPoint.run(async (context) => {
  context.presentation.insertSlidesFromBase64(chosenFileBase64);
  await context.sync();
});
```

## Table Operation Templates

### Create and Fill Table
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // Create 3x4 table
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

### Format Table Cells
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // Create table and apply styles
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

## Shape Operation Templates

### Add Rectangle
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

### Add Text Box
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const textBox = slide.shapes.addTextBox("This is text content");
  textBox.left = 100;
  textBox.top = 100;
  textBox.width = 300;
  textBox.height = 50;
  
  // Set text style
  textBox.textFrame.textRange.font.size = 24;
  textBox.textFrame.textRange.font.color = "purple";
  
  await context.sync();
});
```

### Add Line
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const line = slide.shapes.addLine(PowerPoint.ConnectorType.straight, {
    left: 50,
    top: 50,
    width: 200, // Horizontal distance from start point to end point
    height: 200 // Vertical distance from start point to end point
  });
  
  await context.sync();
});
```

### Group Shapes
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  
  // Assume we want to group first two shapes
  slide.load("shapes/items");
  await context.sync();
  
  if (shapes.items.length >= 2) {
    const shapesToGroup = [shapes.items[0], shapes.items[1]];
    const group = shapes.addGroup(shapesToGroup);
  }
  
  await context.sync();
});
```

## Image Operation Templates

### Add Image
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // base64Image is Base64 string of image (without data:image/... prefix)
  const image = slide.shapes.addImage(base64Image);
  image.left = 100;
  image.top = 100;
  image.width = 400;
  image.height = 300;
  
  await context.sync();
});
```

## Tag (Metadata) Templates

### Add Tag to Slide
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  slide.tags.add("CATEGORY", "Confidential");
  await context.sync();
});
```

### Process Slides by Tag
```javascript
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("tags/key, tags/value");
  await context.sync();
  
  // Iterate and process
  for (let slide of slides.items) {
    // Check tag logic
  }
});
```

## Text Formatting Enhancement Templates

### Set Text Bold and Italic
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // Set text range font properties
  const textRange = shape.textFrame.textRange;
  textRange.font.bold = true;
  textRange.font.italic = true;
  textRange.font.underline = PowerPoint.FontUnderlineStyle.single;
  
  await context.sync();
});
```

### Set Font Name and Size
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

### Vertical Center Text
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // Set text box vertical alignment
  shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
  
  await context.sync();
});
```

### Set Text Box Margins
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // Set text box margins (unit: points)
  shape.textFrame.marginLeft = 10;
  shape.textFrame.marginRight = 10;
  shape.textFrame.marginTop = 5;
  shape.textFrame.marginBottom = 5;
  
  await context.sync();
});
```

## Slide Layout and Master Templates

### Get Slide Layout Information
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // Load layout and master information
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

### Create Slide with Specified Layout
```javascript
PowerPoint.run(async (context) => {
  // Get first slide's layout
  const firstSlide = context.presentation.slides.getItemAt(0);
  firstSlide.load("layout");
  await context.sync();
  
  const layout = firstSlide.layout;
  layout.load("id");
  await context.sync();
  
  // Create new slide with same layout
  const newSlide = context.presentation.slides.add({
    layoutId: layout.id
  });
  
  await context.sync();
});
```

### Get All Available Layouts
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
    
    console.log("Available layouts:");
    for (const layout of layouts.items) {
      layout.load("id,name");
      await context.sync();
      console.log(`- ${layout.name} (ID: ${layout.id})`);
    }
  }
});
```

## Theme System Templates

### Fill Shape with Theme Color
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shape = slide.shapes.getItemAt(0);
  
  // Use theme color (accent1 is theme accent color 1)
  shape.fill.setSolidColor("accent1");
  
  await context.sync();
});
```

### Use Multiple Theme Colors
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  // Create multiple shapes using different theme colors
  const rect1 = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect1.left = 50;
  rect1.top = 50;
  rect1.width = 100;
  rect1.height = 100;
  rect1.fill.setSolidColor("accent1"); // Theme accent color 1
  
  const rect2 = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
  rect2.left = 170;
  rect2.top = 50;
  rect2.width = 100;
  rect2.height = 100;
  rect2.fill.setSolidColor("accent2"); // Theme accent color 2
  
  await context.sync();
});
```

## Table Data Operation Templates

### Read Table Data
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  
  // Find first table
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
    
    // Read all cell data
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

### Update Table Cell
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
    // Update cell content
    const cell = table.getCell(0, 0);
    cell.text = "Updated content";
    
    await context.sync();
  }
});
```

### Set Row Height and Column Width
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
    // Set row height and column width
    const firstRow = table.rows.getItemAt(0);
    firstRow.height = 50;
    
    const firstColumn = table.columns.getItemAt(0);
    firstColumn.width = 150;
    
    await context.sync();
  }
});
```

### Create Table with Merged Cells
```javascript
PowerPoint.run(async (context) => {
  const slide = context.presentation.getSelectedSlides().getItemAt(0);
  
  const tableValues = [
    ["Title", "", "", ""],  // First row will be merged
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

## Presentation Generation Python Templates

### Generate Presentation Outline
```python
def generate_outline(topic, slide_count=5):
    """Generate presentation outline based on topic"""
    prompt = f"""Please generate a {slide_count}-page presentation outline for the following topic:

Topic: {topic}

Please output in the following format:
1. Title Slide: [Title]
2. Content Slide 1: [Title] - [Point1, Point2, Point3]
3. Content Slide 2: [Title] - [Point1, Point2, Point3]
...
{slide_count}. Summary Slide: [Key Conclusions]

Outline:"""
    
    # Call AI to generate outline
    return outline
```

### Generate Slide Content
```python
def generate_slide_content(title, points):
    """Generate detailed content for slide"""
    content = {
        "title": title,
        "bullet_points": points,
        "speaker_notes": generate_speaker_notes(title, points)
    }
    return content

def generate_speaker_notes(title, points):
    """Generate speaker notes"""
    prompt = f"""Please generate concise speaker notes for the following slide:

Title: {title}
Points: {', '.join(points)}

Speaker Notes (50-100 words):"""
    
    # Call AI to generate notes
    return notes
```

### Color Scheme Suggestions
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
    """Get color scheme"""
    return COLOR_SCHEMES.get(style, COLOR_SCHEMES["professional"])
```
