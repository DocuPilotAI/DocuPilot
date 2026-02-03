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

### Example 1: Insert Poem
**User Request**: "Add Li Bai's 'Quiet Night Thoughts' at end of document"

**Your Response**:
```
Alright, I'll insert 'Quiet Night Thoughts' for you.

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

### Example 2: Create Content Control Table
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

## Detailed Templates

For more operation templates, please refer to [TOOLS.md](TOOLS.md).
