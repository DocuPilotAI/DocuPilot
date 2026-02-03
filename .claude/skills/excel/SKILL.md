---
name: excel-operations
description: Excel data operation skill. Used for reading/writing cells, managing worksheets, creating tables and pivot tables, data validation, comments, charts, and data analysis. Use when users mention Excel, tables, cells, charts, pivot tables, or data analysis.
---

# Excel Operations Skill

## Usage

Manipulate Excel by generating **hidden Office.js code** that is automatically executed by the frontend, completely transparent to users.

### Important Rules

- **User Experience First**: Users should only see natural language, not any code
- **Hidden Code Format**: Wrap code with HTML comments: `<!--OFFICE-CODE:excel\ncode\n-->`
- **Friendly Feedback**: Inform users of results in natural language after operations complete
- **Complete & Executable**: Generated code must be complete, directly runnable Office.js code

## ⚠️ API Usage Guidelines & Common Error Patterns

### Understanding Office.js API Structure

The Office.js Excel API follows consistent patterns. Understanding these prevents most common errors:

**1. Collection-Item Pattern**
- Collections are always plural: `worksheets`, `tables`, `charts`, `borders`, `pivotTables`
- Items are singular and accessed via: `getItem(name/id)`, `getItemAt(index)`, `getItemOrNullObject(name/id)`
- Rule: Never call collection methods on items, or item methods on collections

**2. Format Object Hierarchy**
- Format properties are nested: `range.format.{category}.{property}`
- Categories: `font`, `fill`, `borders`, `protection`, alignment properties
- Border operations require the **`borders`** collection (plural), not `border` (singular)

**3. Array-Based Properties**
- Properties like `values`, `formulas`, `numberFormat` always expect 2D arrays
- Rule: Array dimensions must exactly match range dimensions
- Even single cells require 2D array structure: `[[value]]`

### Multi-Range Operations (RangeAreas)

- ❌ **Wrong**: `sheet.getRangeAreas("B3,F3,J3")` - Worksheet object doesn't have getRangeAreas method
- ✅ **Correct**: `workbook.getRangeAreas("Sheet1!B3,Sheet1!F3,Sheet1!J3")` - Must use Workbook object and include sheet name
- ✅ **Recommended**: Operate each cell separately: `sheet.getRange("B3")`, `sheet.getRange("F3")`, etc.

**When to use RangeAreas:**
- Cross-worksheet operations
- Batch formatting operations on non-contiguous ranges
- Limited operation support compared to single Range objects

**When to use separate Range calls:**
- Setting different values for each range
- Complex operations on each range
- More readable and maintainable code

### Error Patterns and Prevention

**Pattern 1: InvalidArgument - Invalid or missing parameter**
- **Causes**: Non-existent object reference, wrong parameter type, value out of range
- **Prevention**: Use `getItemOrNullObject()` and check `isNullObject` before operations
```javascript
const sheet = context.workbook.worksheets.getItemOrNullObject("SheetName");
await context.sync();
if (sheet.isNullObject) {
  console.log("Worksheet doesn't exist");
  return;
}
```

**Pattern 2: InvalidReference - Invalid cell/range reference**
- **Causes**: Referenced cell or range doesn't exist in worksheet
- **Prevention**: Validate range address format and bounds before accessing
```javascript
const range = sheet.getRange("A1:Z1000");
range.load("address");
await context.sync();
// Range is now validated and safe to use
```

**Pattern 3: Array Dimension Mismatch**
- **Error Message**: "The number of rows or columns in the input array doesn't match the size or dimensions of the range"
- **Causes**: Array size doesn't match range size when setting `values`, `formulas`, or `numberFormat`
- **Prevention**: Calculate array dimensions based on range size
```javascript
// For range "A1:B2" (2 rows × 2 columns)
range.values = [[1, 2], [3, 4]]; // ✅ Correct: 2×2 array

// For single cell "A1" (1 row × 1 column)
range.values = [[100]]; // ✅ Correct: 1×1 array
```

**Pattern 4: Property/Method Not Found**
- **Error Message**: "undefined is not an object" or "is not a function"
- **Causes**: 
  - Using singular instead of plural (e.g., `border` instead of `borders`)
  - Calling collection methods on items or vice versa
  - Accessing properties before `load()` and `sync()`
- **Prevention**: Follow API hierarchy, use correct collection names, load properties before access

**Pattern 5: Requested Resource Does Not Exist**
- **Causes**: Attempting to delete or modify non-existent objects
- **Prevention**: Always use `getItemOrNullObject()` before delete operations
```javascript
const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
await context.sync();
if (!pivotTable.isNullObject) {
  pivotTable.delete();
  await context.sync();
}
```

### Load and Sync Best Practices

1. **Load Pattern**: Must load properties before reading them
   ```javascript
   range.load("values, address, formulas");
   await context.sync();
   console.log(range.values); // Now safe to access
   ```

2. **Write Pattern**: Set properties then sync to commit changes
   ```javascript
   range.values = [[100]];
   range.format.font.bold = true;
   await context.sync(); // Commits all changes
   ```

3. **Batch Operations**: Minimize `sync()` calls for better performance
   ```javascript
   // ✅ Good: Single sync for multiple operations
   range1.values = [[1]];
   range2.values = [[2]];
   range3.values = [[3]];
   await context.sync();
   
   // ❌ Bad: Multiple syncs
   range1.values = [[1]];
   await context.sync();
   range2.values = [[2]];
   await context.sync();
   ```

4. **Cross-Context Access**: Properties loaded in one context can't be accessed in another
   ```javascript
   // ❌ Wrong: Trying to use loaded data outside Excel.run
   let data;
   await Excel.run(async (context) => {
     const range = sheet.getRange("A1");
     range.load("values");
     await context.sync();
     data = range.values; // Store before context ends
   });
   console.log(data); // ✅ Now accessible outside context
   ```

### Async Operations and Context Management

1. **Context Scope**: Each `Excel.run()` creates an isolated context
   - Objects created in one context cannot be directly used in another
   - Store primitive values (strings, numbers, arrays) to pass between contexts

2. **Exception Handling**: Wrap operations in try-catch for robust error handling
   ```javascript
   try {
     await Excel.run(async (context) => {
       // Operations here
       await context.sync();
     });
   } catch (error) {
     console.error("Excel operation failed:", error);
     // Provide user-friendly error message
   }
   ```

3. **Async/Await Requirements**:
   - All `context.sync()` calls must be awaited
   - Excel.run callback must be async function
   - Don't use callbacks or promises without proper async/await handling

### Enum Constants and Type Safety

Always use Excel namespace enums for better code quality:

```javascript
// ✅ Correct: Using enums
range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
border.style = Excel.BorderLineStyle.continuous;
border.weight = Excel.BorderWeight.thick;

// ❌ Avoid: Using strings (error-prone)
range.format.horizontalAlignment = "center"; // May work but not type-safe
border.style = "Continuous"; // Case-sensitive, easy to typo
```

Common enum namespaces:
- `Excel.BorderIndex`: edgeTop, edgeBottom, edgeLeft, edgeRight, insideHorizontal, insideVertical
- `Excel.BorderLineStyle`: none, continuous, dash, dashDot, dot, double
- `Excel.BorderWeight`: hairline, thin, medium, thick
- `Excel.HorizontalAlignment`: left, center, right, justify, distributed
- `Excel.VerticalAlignment`: top, center, bottom, justify, distributed
- `Excel.ChartType`: columnClustered, line, pie, bar, area, scatter, etc.
- `Excel.RangeUnderlineStyle`: none, single, double, singleAccountant, doubleAccountant

### Parameter Validation Rules

1. **Range Addresses**: 
   - Must follow A1 notation: "A1", "B2:D10", "Sheet1!A1:B5"
   - Column letters are case-insensitive: "A1" equals "a1"
   - Row numbers start from 1, not 0

2. **Array Data**:
   - Outer array represents rows, inner arrays represent columns
   - All inner arrays must have same length
   - Empty cells should be empty string "" or null, not undefined

3. **Object Names**:
   - Worksheet names: avoid special characters `[]/*?:`
   - Table names: must start with letter or underscore, no spaces
   - Named ranges: similar to table names, follow Excel naming rules

### Performance Optimization Guidelines

1. **Minimize Sync Calls**: Each sync is a round-trip to Excel
   ```javascript
   // ✅ Efficient: 1 sync for 100 operations
   for (let i = 0; i < 100; i++) {
     sheet.getRange(`A${i+1}`).values = [[i]];
   }
   await context.sync();
   ```

2. **Use Bulk Operations**: Prefer range operations over cell-by-cell
   ```javascript
   // ✅ Better: Single range operation
   range.values = [[1,2,3], [4,5,6], [7,8,9]];
   
   // ❌ Slower: Multiple cell operations
   sheet.getRange("A1").values = [[1]];
   sheet.getRange("B1").values = [[2]];
   // ... 7 more operations
   ```

3. **Load Only Required Properties**: Minimize data transfer
   ```javascript
   // ✅ Efficient: Load specific properties
   range.load("values, address");
   
   // ❌ Inefficient: Load everything
   range.load();
   ```

4. **Batch Similar Operations**: Group similar operations together
   ```javascript
   // Set all formatting first, then all values
   range1.format.font.bold = true;
   range2.format.font.bold = true;
   range3.format.font.bold = true;
   range1.values = [[1]];
   range2.values = [[2]];
   range3.values = [[3]];
   await context.sync();
   ```

### Feature-Specific API Guidelines

#### Conditional Formatting
- Each range can apply multiple conditional formats, evaluated in priority order
- Use `conditionalFormat.priority` to adjust evaluation order
- Use `conditionalFormat.stopIfTrue` to prevent subsequent rule evaluation
- Icon set `criteria` array: index 0 is lowest level, last index is highest level
- Format options are accessed via nested objects: `conditionalFormat.{type}.format.{category}`

#### Event Handling
- Event handlers are destroyed when add-in refreshes or closes
- Pattern: save `EventResult` object → call `eventResult.remove()` to cleanup
- Temporarily disable events during batch operations: `context.runtime.enableEvents = false`
- Events are worksheet or workbook-level, accessed via `sheet.on{EventName}.add(handler)`

#### Shape Operations
- Position properties (`left`, `top`) use points, not pixels
- Image insertion requires Base64 string without `data:image/png;base64,` prefix
- Shape grouping: load shape `id` property first via `shape.load("id")`
- Name shapes for later reference: `shape.name = "MyShape"`
- Shape collections are accessed via `sheet.shapes` or `chart.shapes`

#### Notes vs Comments
- **Notes**: Traditional yellow sticky notes (one per cell)
- **Comments**: Threaded discussion comments (multiple per cell)
- Access via: `worksheet.notes` or `workbook.notes` for notes collection
- Property `note.visible` controls permanent visibility (default: hover only)
- Use `notes.add(cellAddress, content)` to create notes

#### Range Advanced Operations
- `copyFrom(source, copyType, skipBlanks, transpose)`: skipBlanks preserves destination data
- `moveTo(destinationRange)`: cut-paste operation, auto-expands destination
- `removeDuplicates(columns, includesHeader)`: column indices are 0-based
- `insert(shift)` and `delete(shift)`: affect surrounding ranges in worksheet
- Always consider impact on other ranges when using insert/delete

#### Checkboxes
- Cell control type for boolean value visualization
- Conversion: `range.control = { type: Excel.CellControlType.checkbox }`
- State management: use `range.values` with `[[true]]` or `[[false]]`
- Read state: `range.load("values")` then `range.values` returns boolean array
- Removal: `range.control = { type: Excel.CellControlType.empty }`

## Workflow

1. **Understand Requirements**: Analyze user's data operation requests
2. **Reference Template**: Check code templates in TOOLS.md
3. **Generate Code**: Create complete Office.js code
4. **Embed Hidden Marker**: Wrap code with `<!--OFFICE-CODE:excel ... -->`
5. **Add Friendly Message**: Inform user of operation results

## Supported Features

- **Worksheet Management**: Create, rename, delete, copy, activate worksheets.
  - ⚠️ Must use `getItemOrNullObject` to check object existence before delete operations
- **Data Read/Write**: Read/write cells, ranges, array data.
- **Table Operations**: Create tables, sort, filter, add rows/columns.
- **Pivot Tables** (Full Support):
  - Create pivot tables, add row/column/data/filter fields
  - Apply filters (date filter, label filter, value filter)
  - Create and use slicers for interactive filtering
  - Switch layout types (compact, outline, tabular)
  - Get and format pivot table data
  - Refresh and delete pivot tables
  - ⚠️ Must check pivot table existence before all operations
- **Conditional Formatting** (Full Support):
  - Cell value conditions (cellValue): Apply format based on cell values
  - Data bars (dataBar): Display data bars in cells
  - Icon sets (iconSet): Visualize data with arrows, icons, etc.
  - Preset conditions (preset): Preset rules like above/below average
  - Top/Bottom N (topBottom): Highlight top N or bottom N
  - Custom formula (custom): Set conditions using custom formulas
- **Event Handling**:
  - Data change event (onChanged): Monitor cell data changes
  - Selection change event (onSelectionChanged): Monitor user selection changes
  - Worksheet activation event (onActivated): Monitor worksheet switches
  - Calculation complete event (onCalculated): Monitor worksheet calculation completion
  - ⚠️ Event handlers need proper registration and removal to avoid memory leaks
- **Shape Operations**:
  - Geometric shapes: Add rectangles, circles, arrows, etc.
  - Image insertion: Support Base64 encoded JPEG/PNG images
  - Lines and connectors: Create straight lines, arrow lines, etc.
  - Text boxes: Create and format text boxes
  - Shape grouping: Combine multiple shapes into group
- **Notes**:
  - Add cell notes (traditional yellow sticky notes)
  - Modify note content and visibility
  - Delete notes
  - ⚠️ Different from Comments (threaded comments), Notes are traditional note feature
- **Range Advanced Operations**:
  - Copy paste: Support skip blanks, transpose options
  - Move range: Cut and paste to new location
  - Insert/delete cells: Specify shift direction
  - Clear content: Clear values, formats, or all
  - Remove duplicates: Remove duplicate rows based on specified columns
  - Row/column grouping: Create outline groups and ungroup
- **Checkboxes**:
  - Convert boolean values to checkbox controls
  - Read checkbox state
  - Remove checkbox to restore to boolean value
- **Data Validation**: Set validation rules like numeric ranges, date limits, dropdown lists.
- **Comments & Naming**: Add comments, create named ranges.
- **Charts**: Create various chart types (column charts, line charts, etc.).
- **Formatting**: Set number formats, fonts, colors.

## Pivot Table Operation Notes

### Prerequisite Checks
Must check object existence before pivot table operations to avoid "Requested resource does not exist" errors:

```javascript
const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
await context.sync();

if (pivotTable.isNullObject) {
  console.log("Pivot table doesn't exist");
  return;
}
```

### Field References
Pivot table fields must come from source data column names, ensure:
1. Field name exactly matches source data column header
2. Use `getItemOrNullObject` to check field existence
3. Sync before operating fields to ensure fields are loaded

### Filter Operations
- **Date Filter**: Need to specify date condition and precision (year/month/day)
- **Label Filter**: For text fields, supports conditions like contains/excludes, starts with/ends with
- **Value Filter**: Filter by aggregated values, need to specify numeric field and comparison condition

### Slicer Creation
Slicers depend on pivot tables, before creation must:
1. Ensure pivot table exists
2. Ensure filter field exists in pivot table
3. Set reasonable position and size to avoid overlap

## Examples

### Example 1: Manage Worksheets
**User Request**: "Create a new worksheet named 'Sales Report'"

**Your Response**:
```
Alright, creating a new worksheet for you.

<!--OFFICE-CODE:excel
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.add("Sales Report");
  sheet.activate();
  await context.sync();
});
-->

✅ Successfully created and activated worksheet named "Sales Report".
```

### Example 2: Create Pivot Table
**User Request**: "Create a pivot table with current data, summarize sales by product"

**Your Response**:
```
Alright, I'll create a pivot table for you.

<!--OFFICE-CODE:excel
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getUsedRange();
  
  const pivotTable = sheet.pivotTables.add("PivotTable1", range, sheet.getRange("F1"));
  
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Product"));
  pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Sales"));
  
  await context.sync();
});
-->

✅ Created pivot table, summarized sales data by product, starting position at F1 cell.
```

### Example 3: Set Data Validation
**User Request**: "Limit column B to only accept integers between 1 and 100"

**Your Response**:
```
Alright, I'll set data validation rules for column B.

<!--OFFICE-CODE:excel
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B:B");
  
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: 1,
      formula2: 100,
      operator: Excel.DataValidationOperator.between
    }
  };
  
  range.dataValidation.errorAlert = {
    showAlert: true,
    title: "Input Error",
    message: "Please enter an integer between 1 and 100",
    style: Excel.DataValidationAlertStyle.stop
  };
  
  await context.sync();
});
-->

✅ Set data validation for column B: only allows integers from 1-100.
```

## User File Handling

### Workspace File Access

Users can upload files to current session, these files are saved in workspace directory:

- **User Upload File Path**: `workspace/sessions/{session_id}/uploads/`
- **Generated File Save Path**: `workspace/sessions/{session_id}/outputs/`

### File Operation Flow

1. **Find User-Uploaded Files**:
   ```typescript
   // Use Glob tool to find Excel files
   // Filenames include timestamp prefix, use wildcards
   const pattern = "workspace/sessions/{session_id}/uploads/*.xlsx";
   ```

2. **Read File Data**:
   - For text format files (CSV, TXT, JSON), use Read tool
   - For Excel files, guide user to open in Excel then use Office.js API to operate

3. **Save Analysis Results**:
   ```typescript
   // Use Write tool to save analysis report
   Write: workspace/sessions/{session_id}/outputs/analysis_report.txt
   ```

### Example Workflow

**User Request**: "Analyze my uploaded sales data table"

**Processing Steps**:
1. Use Glob to find: `workspace/sessions/abc123/uploads/*.xlsx`
2. Guide user: "I found your uploaded file `sales_data.xlsx`. Please open this file in Excel, then I can help you analyze the data."
3. After user opens file in Excel, use `office_excel_*` tools to read and analyze data
4. Save analysis results to: `workspace/sessions/abc123/outputs/sales_analysis.txt`

## Detailed Templates

For more operation templates, please refer to [TOOLS.md](TOOLS.md).
