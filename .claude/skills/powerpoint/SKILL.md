---
name: powerpoint-operations
description: PowerPoint presentation operation skill. Used for creating slides, adding content (text, images, shapes, tables), formatting, and generating presentation outlines. Use when users mention PPT, slides, presentations, or speeches.
---

# PowerPoint Operations Skill

## Usage

Manipulate PowerPoint by generating **hidden Office.js code** that is automatically executed by the frontend, completely transparent to users.

### Important Rules

- **User Experience First**: Users should only see natural language, not any code
- **Hidden Code Format**: Wrap code with HTML comments: `<!--OFFICE-CODE:powerpoint\ncode\n-->`
- **Friendly Feedback**: Inform users of results in natural language after operations complete
- **Complete & Executable**: Generated code must be complete, directly runnable Office.js code

## Workflow

1. **Understand Requirements**: Analyze user's presentation operation requests
2. **Reference Template**: Check code templates in TOOLS.md
3. **Generate Code**: Create complete Office.js code
4. **Embed Hidden Marker**: Wrap code with `<!--OFFICE-CODE:powerpoint ... -->`
5. **Add Friendly Message**: Inform user of operation results

## Supported Features

- **Slide Management**: Add, delete, insert (from Base64), get list.
- **Content Addition**:
  - **Text Box**: Set text, font, color, alignment.
  - **Shapes**: Add geometric shapes (rectangles, circles, etc.), lines, connectors.
  - **Images**: Insert Base64 images, set position and size.
  - **Tables**: Create tables, fill data, set styles (borders, fill, fonts).
- **Formatting**: Adjust position, size, fill color, line style.
- **Group/Ungroup**: Manage shape groups.
- **Metadata**: Add and read tags.
- **Text Formatting Enhancement**:
  - Complete font properties: bold, italic, underline, font name, size, color
  - Text box vertical alignment: top, center, bottom
  - Text box margin settings: left, right, top, bottom margin control
  - Text auto-fit: automatically shrink to fit shape
- **Slide Layouts and Masters**:
  - Get current slide's layout and master information
  - Create new slides using specified layout
  - List all available slide layouts
  - Query master name and ID
- **Theme System**:
  - Use theme colors to fill shapes (accent1-6, background, text, etc.)
  - Theme colors automatically adapt to presentation theme
  - Supports 12 standard theme colors
- **Table Data Operations**:
  - Read all table cell data
  - Update specified cell content
  - Set row height and column width
  - Create tables with merged cells
  - Format table cells (background, font, borders)

## ⚠️ Common Error Handling

### InvalidArgument Error
- **Cause**: Slide index out of range, invalid shape ID
- **Solution**: Ensure index is valid when using `getItemAt()`
```javascript
const slides = context.presentation.slides;
slides.load("items");
await context.sync();

if (index >= 0 && index < slides.items.length) {
  const slide = slides.getItemAt(index);
  // Operate slide
}
```

### InvalidReference Error
- **Cause**: Referencing deleted slide or shape
- **Solution**: Don't continue referencing object after deletion

### Text Formatting Notes
- All font properties are on `textFrame.textRange.font` object
- Vertical alignment is on `textFrame.verticalAlignment`
- Margin units are points, not pixels
- Use `FontUnderlineStyle` enum to set underline style

### Theme Color Usage
- Theme color names: accent1-accent6, background1-background2, text1-text2, hyperlink, followedHyperlink
- Using theme colors ensures consistent presentation style
- Shapes using theme colors will automatically update if theme changes

### Table Operation Limitations
- Table cell indices start from 0
- Reading cell data requires multiple syncs, note performance
- Merged cells are specified through `mergedAreas` parameter when creating table
- Table's `shape.type` property value is `PowerPoint.ShapeType.table`

### Best Practices
1. Load and sync before accessing collections
2. Use reasonable numeric ranges for coordinates and dimensions
3. Use tags to manage slide metadata
4. Text box size should be sufficient to contain content, avoid overflow
5. Using layouts and masters can maintain consistent presentation style
6. **CRITICAL**: When adding a new slide with `slides.add()`, you MUST:
   - Call `await context.sync()` immediately after `add()`
   - Load the slides collection and sync again to get a reference to the new slide
   - Only then can you add shapes to the new slide
   
   ```javascript
   // ✅ CORRECT: Add slide, sync, get reference, then add shapes
   context.presentation.slides.add();
   await context.sync();
   
   const slides = context.presentation.slides;
   slides.load("items");
   await context.sync();
   const newSlide = slides.items[slides.items.length - 1];
   
   // Now safe to add shapes
   const textBox = newSlide.shapes.addTextBox("text");
   ```
   
   ```javascript
   // ❌ WRONG: Try to use slide immediately after add()
   const slide = context.presentation.slides.add();
   const textBox = slide.shapes.addTextBox("text"); // May fail silently!
   ```

## Examples

### Example 1: Create Title Slide
**User Request**: "Create a slide with title '2023 Annual Report'"

**Your Response**:
```
Alright, I've created a title slide for you.

<!--OFFICE-CODE:powerpoint
PowerPoint.run(async (context) => {
  // Add new slide
  context.presentation.slides.add();
  await context.sync();
  
  const slide = context.presentation.slides.getItemAt(context.presentation.slides.items.length - 1);
  
  // Add title text box
  const textBox = slide.shapes.addTextBox("2023 Annual Report");
  textBox.left = 100;
  textBox.top = 200;
  textBox.width = 760;
  textBox.height = 100;
  
  // Set title style
  textBox.textFrame.textRange.font.size = 60;
  textBox.textFrame.textRange.font.bold = true;
  textBox.textFrame.textRange.paragraphFormat.align = PowerPoint.ParagraphAlignment.center;
  
  await context.sync();
});
-->

✅ Created new slide with title "2023 Annual Report".
```

### Example 2: Insert Data Table
**User Request**: "Insert a 3x3 data table in current slide"

**Your Response**:
```
Alright, inserting table in current slide.

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

✅ Inserted 3x3 data table with blue style applied.
```

## User File Handling

### Workspace File Access

Users can upload files (like images) to current session, these files are saved in workspace directory:

- **User Upload File Path**: `workspace/sessions/{session_id}/uploads/`
- **Generated File Save Path**: `workspace/sessions/{session_id}/outputs/`

### File Operation Flow

1. **Find User-Uploaded Images**:
   ```typescript
   // Use Glob tool to find image files
   // Filenames include timestamp prefix, use wildcards
   const pattern = "workspace/sessions/{session_id}/uploads/*.{png,jpg,jpeg}";
   ```

2. **Process Images**:
   - PowerPoint image insertion requires Base64 encoding
   - If user uploads images, need to read file and convert to Base64 first

3. **Save Presentation Outline**:
   ```typescript
   // Use Write tool to save presentation outline
   Write: workspace/sessions/{session_id}/outputs/presentation_outline.txt
   ```

### Example Workflow

**User Request**: "Create presentation using my uploaded images"

**Processing Steps**:
1. Use Glob to find: `workspace/sessions/abc123/uploads/*.png`
2. Inform user: "I found your uploaded image `logo.png`. I'll create a presentation and insert this image for you."
3. Read image file, convert to Base64 (if needed)
4. Use Office.js API to create slides and insert images
5. If need to save presentation outline, save to: `workspace/sessions/abc123/outputs/slides_outline.txt`

## Detailed Templates

For more operation templates, please refer to [TOOLS.md](TOOLS.md).
