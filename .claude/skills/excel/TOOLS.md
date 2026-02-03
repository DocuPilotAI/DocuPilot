# Excel Tool Template Library

## Worksheet Management Templates

### Create Worksheet
```javascript
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.add("NewSheet");
  sheet.activate();
  await context.sync();
});
```

### Create Worksheet (With Validation)
```javascript
// ⚠️ Best Practice: Check if worksheet already exists
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const existingSheet = sheets.getItemOrNullObject("NewSheet");
  await context.sync();
  
  let sheet;
  if (existingSheet.isNullObject) {
    // Worksheet doesn't exist, create new one
    sheet = sheets.add("NewSheet");
  } else {
    // Worksheet exists, use existing one
    sheet = existingSheet;
  }
  
  sheet.activate();
  await context.sync();
});
```

### Rename Worksheet
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.name = "RenamedSheet";
  await context.sync();
});
```

### Delete Worksheet
```javascript
// ⚠️ Important: Check if worksheet exists before deletion, recommend using "Delete Worksheet (With Validation)" template
try {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("SheetToDelete");
    sheet.delete();
    await context.sync();
  });
} catch (error) {
  console.error("Error deleting worksheet:", error);
  throw error;
}
```

### Delete Worksheet (With Validation)
```javascript
// ⚠️ Best Practice: Check worksheet exists before deletion (Recommended)
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject("SheetToDelete");
  await context.sync();
  
  if (!sheet.isNullObject) {
    sheet.delete();
    await context.sync();
    console.log("Worksheet deleted");
  } else {
    console.log("Worksheet doesn't exist, no need to delete");
  }
});
```

### Copy Worksheet
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const newSheet = sheet.copy(Excel.WorksheetPositionType.after, sheet);
  newSheet.name = "CopiedSheet";
  await context.sync();
});
```

## Data Reading Templates

### Read Selected Range
```javascript
// Get currently selected cell range
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  range.load(["values", "address", "formulas", "numberFormat"]);
  await context.sync();
  
  return {
    address: range.address,
    values: range.values,
    formulas: range.formulas,
    format: range.numberFormat
  };
});
```

### Read Specific Range
```javascript
// Read cells at specific address
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  range.load("values");
  await context.sync();
  
  return range.values;
});
```

### Read Multiple Non-contiguous Ranges (Important!)
```javascript
// ⚠️ Note: getRangeAreas is a workbook-level method, not sheet-level
// Wrong usage: sheet.getRangeAreas("B3,F3,J3") ❌
// Correct usage: workbook.getRangeAreas() ✅

Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sheetName = sheet.name;
  
  // Use workbook.getRangeAreas() to read multiple non-contiguous ranges
  const rangeAreas = context.workbook.getRangeAreas(`${sheetName}!B3,${sheetName}!F3,${sheetName}!J3`);
  rangeAreas.load("address");
  
  // Or read each cell separately
  const range1 = sheet.getRange("B3");
  const range2 = sheet.getRange("F3");
  const range3 = sheet.getRange("J3");
  
  range1.load("values");
  range2.load("values");
  range3.load("values");
  
  await context.sync();
  
  return {
    b3: range1.values[0][0],
    f3: range2.values[0][0],
    j3: range3.values[0][0]
  };
});
```

## Data Writing Templates

### Write Single Value
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const cell = sheet.getRange("A1");
  cell.values = [["Hello World"]];
  await context.sync();
});
```

### Write Array Data
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:C3");
  range.values = [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9]
  ];
  await context.sync();
});
```

### Write to Multiple Non-contiguous Ranges
```javascript
// Method 1: Write each range separately (Recommended)
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.getRange("B3").values = [[10]];
  sheet.getRange("F3").values = [[20]];
  sheet.getRange("J3").values = [[30]];
  
  await context.sync();
});

// Method 2: Use workbook.getRangeAreas() (Cross-sheet scenarios)
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sheetName = sheet.name;
  
  // Note: Must use complete address (including sheet name)
  const rangeAreas = context.workbook.getRangeAreas(
    `${sheetName}!B3,${sheetName}!F3,${sheetName}!J3`
  );
  
  // RangeAreas has limited operations, usually used for formatting and batch operations
  rangeAreas.format.fill.color = "yellow";
  
  await context.sync();
});
```

## Table Operation Templates

### Create Table
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  // Create table with header row
  const table = sheet.tables.add("A1:D5", true);
  table.name = "SalesTable";
  
  // Set headers
  table.getHeaderRowRange().values = [["Date", "Product", "Category", "Amount"]];
  
  // Set data
  table.getDataBodyRange().values = [
    ["2023-01-01", "Widget A", "Electronics", 100],
    ["2023-01-02", "Widget B", "Electronics", 200],
    ["2023-01-03", "Widget C", "Home", 150],
    ["2023-01-04", "Widget D", "Home", 300]
  ];
  
  // Auto-fit columns
  sheet.getUsedRange().format.autofitColumns();
  
  await context.sync();
});
```

### Sort Table
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItem("SalesTable");
  
  // Sort by 4th column (Amount) in descending order
  table.sort.apply([
    {
      key: 3, // 4th column, index starts from 0
      ascending: false
    }
  ]);
  
  await context.sync();
});
```

### Filter Table
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItem("SalesTable");
  
  // Filter 3rd column (Category) for "Electronics"
  table.columns.getItemAt(2).filter.apply({
    filterOn: Excel.FilterOn.values,
    values: ["Electronics"]
  });
  
  await context.sync();
});
```

## Pivot Table Templates

### Create Pivot Table
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  // Assume data is in A1:D5
  const sourceRange = sheet.getRange("A1:D5");
  
  // Create pivot table at F1
  const pivotTable = sheet.pivotTables.add("PivotTable1", sourceRange, "F1");
  
  // Add row fields
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Category"));
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Product"));
  
  // Add data field (default sum)
  pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Amount"));
  
  await context.sync();
});
```

### Add Column Hierarchy and Filters
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Add column field
  pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Region"));
  
  // Add filter field
  pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Category"));
  
  await context.sync();
});
```

### Pivot Table Date Filter
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Get or add date hierarchy
  let dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date");
  await context.sync();
  
  if (dateHierarchy.isNullObject) {
    dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date"));
    await context.sync();
  }
  
  // Apply date filter: only show data after 2020-08-01
  const filterField = dateHierarchy.fields.getItem("Date");
  const dateFilter = {
    condition: Excel.DateFilterCondition.afterOrEqualTo,
    comparator: {
      date: "2020-08-01",
      specificity: Excel.FilterDatetimeSpecificity.month
    }
  };
  filterField.applyFilter({ dateFilter: dateFilter });
  
  await context.sync();
});
```

### Pivot Table Label Filter
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Get field
  const filterField = pivotTable.rowHierarchies.getItem("Category")
    .fields.getItem("Category");
  
  // Label filter: exclude items starting with "Electronics"
  const labelFilter = {
    condition: Excel.LabelFilterCondition.beginsWith,
    substring: "Electronics",
    exclusive: true
  };
  filterField.applyFilter({ labelFilter: labelFilter });
  
  await context.sync();
});
```

### Pivot Table Value Filter
```javascript
// ⚠️ Value filter: Filter row field items by aggregated values
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
  await context.sync();
  
  if (pivotTable.isNullObject) {
    console.log("Pivot table doesn't exist");
    return;
  }
  
  // Get field in row hierarchy
  const productHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Product");
  await context.sync();
  
  if (!productHierarchy.isNullObject) {
    const filterField = productHierarchy.fields.getItem("Product");
    
    // Value filter: only show items with sales greater than 500
    const valueFilter = {
      condition: Excel.ValueFilterCondition.greaterThan,
      comparator: 500,
      value: "Amount"
    };
    filterField.applyFilter({ valueFilter: valueFilter });
  }
  
  await context.sync();
});
```

### Clear Pivot Table Filters
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Load all hierarchies
  pivotTable.hierarchies.load("items");
  await context.sync();
  
  // Clear filters from all fields
  for (const hierarchy of pivotTable.hierarchies.items) {
    hierarchy.fields.load("items");
    await context.sync();
    
    for (const field of hierarchy.fields.items) {
      field.clearAllFilters();
    }
  }
  
  await context.sync();
});
```

### Create Slicer
```javascript
// ⚠️ Create slicer for interactive pivot table filtering
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // First check if pivot table exists
  const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
  await context.sync();
  
  if (pivotTable.isNullObject) {
    console.log("Pivot table doesn't exist, cannot create slicer");
    return;
  }
  
  // Create slicer for interactive filtering
  const slicer = sheet.slicers.add(
    "PivotTable1",  // Pivot table name
    "Category"      // Filter field
  );
  
  slicer.name = "Category Slicer";
  slicer.left = 400;
  slicer.top = 200;
  slicer.width = 200;
  slicer.height = 200;
  
  await context.sync();
});
```

### Filter with Slicer
```javascript
Excel.run(async (context) => {
  const slicer = context.workbook.slicers.getItem("Category Slicer");
  
  // Select specific items for filtering
  slicer.selectItems(["Electronics", "Home", "Clothing"]);
  
  await context.sync();
});
```

### Switch Pivot Table Layout
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Switch layout type: Compact, Outline, Tabular
  pivotTable.layout.layoutType = "Outline";  // or "Compact", "Tabular"
  
  await context.sync();
});
```

### Get Pivot Table Data
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Get data range
  const dataRange = pivotTable.layout.getDataBodyRange();
  dataRange.load("address, values");
  
  await context.sync();
  
  console.log("Data range:", dataRange.address);
  console.log("Data values:", dataRange.values);
  
  return dataRange.values;
});
```

### Format Pivot Table
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  const pivotLayout = pivotTable.layout;
  
  // Set empty cell display text
  pivotLayout.emptyCellText = "--";
  pivotLayout.fillEmptyCells = true;
  
  // Preserve formatting settings
  pivotLayout.preserveFormatting = true;
  
  // Set data range alignment
  const dataRange = pivotLayout.getDataBodyRange();
  dataRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;
  
  await context.sync();
});
```

### Refresh Pivot Table
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Refresh pivot table data
  pivotTable.refresh();
  
  await context.sync();
});
```

### Delete Pivot Table
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // Delete pivot table
  pivotTable.delete();
  
  await context.sync();
});
```

## Data Validation Templates

### Set Numeric Validation
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // Only allow integers greater than 0
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: 0,
      operator: Excel.DataValidationOperator.greaterThan
    }
  };
  
  // Set input prompt
  range.dataValidation.prompt = {
    showPrompt: true,
    title: "Input Limitation",
    message: "Please enter an integer greater than 0"
  };
  
  // Set error alert
  range.dataValidation.errorAlert = {
    showAlert: true,
    style: Excel.DataValidationAlertStyle.stop,
    title: "Invalid Input",
    message: "Value must be greater than 0"
  };
  
  await context.sync();
});
```

### Set Dropdown List
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("C2:C10");
  
  range.dataValidation.rule = {
    list: {
      inCellDropDown: true,
      source: "Option1,Option2,Option3"
    }
  };
  
  await context.sync();
});
```

## Comments & Named Items

### Add Comment
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const comments = sheet.comments;
  
  // Add comment at A1
  comments.add("A1", "This is a comment.");
  
  await context.sync();
});
```

### Create Named Range
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:B10");
  
  // Create named range "MyData"
  context.workbook.names.add("MyData", range);
  
  await context.sync();
});
```

## Chart Creation Templates

### Column Chart
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const dataRange = sheet.getRange("A1:B10");
  
  const chart = sheet.charts.add(
    Excel.ChartType.columnClustered,
    dataRange,
    Excel.ChartSeriesBy.auto
  );
  
  chart.title.text = "Data Analysis Chart";
  chart.setPosition("E1", "L15");
  
  await context.sync();
});
```

### Line Chart
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const dataRange = sheet.getRange("A1:B10");
  
  const chart = sheet.charts.add(
    Excel.ChartType.line,
    dataRange,
    Excel.ChartSeriesBy.columns
  );
  
  chart.title.text = "Trend Analysis";
  await context.sync();
});
```

## Formatting Templates

### Font Formatting
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // Font properties are under format.font
  range.format.font.name = "Arial";
  range.format.font.size = 12;
  range.format.font.bold = true;
  range.format.font.italic = false;
  range.format.font.underline = Excel.RangeUnderlineStyle.single;
  range.format.font.color = "#FF0000";
  
  await context.sync();
});
```

### Fill Color
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // Fill color is under format.fill
  range.format.fill.color = "#FFFF00"; // Yellow background
  
  await context.sync();
});
```

### Number Format
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // Number format uses 2D array structure
  range.numberFormat = [["#,##0.00"]]; // Thousands separator, two decimal places
  // Other examples:
  // range.numberFormat = [["0.00%"]]; // Percentage
  // range.numberFormat = [["$#,##0.00"]]; // Currency
  // range.numberFormat = [["m/d/yyyy"]]; // Date format
  
  await context.sync();
});
```

### Alignment
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // Alignment properties are directly under format
  range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
  range.format.wrapText = true;
  range.format.textOrientation = 0; // 0-90 degrees
  
  await context.sync();
});
```

### Border Formatting
```javascript
// ⚠️ API Rule: Use format.borders (plural collection), not format.border (doesn't exist)
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // Available BorderIndex values:
  // - edgeTop, edgeBottom, edgeLeft, edgeRight (outer borders)
  // - insideHorizontal, insideVertical (inner borders)
  // - diagonalDown, diagonalUp (diagonal lines)
  
  // Set outer borders
  range.format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeLeft).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeRight).style = Excel.BorderLineStyle.continuous;
  
  // Set inner borders (only for multi-cell ranges)
  range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.insideVertical).style = Excel.BorderLineStyle.continuous;
  
  // Customize border properties
  const topBorder = range.format.borders.getItem(Excel.BorderIndex.edgeTop);
  topBorder.color = "#000000";
  topBorder.weight = Excel.BorderWeight.thick; // thin, medium, thick, hairline
  
  await context.sync();
});
```

### Border Formatting (All Borders Loop)
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // Apply same style to all borders using loop
  const borderIndices = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical
  ];
  
  for (const index of borderIndices) {
    const border = range.format.borders.getItem(index);
    border.style = Excel.BorderLineStyle.continuous;
    border.color = "#000000";
    border.weight = Excel.BorderWeight.thin;
  }
  
  await context.sync();
});
```

### Clear Borders
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // Clear all borders by setting style to none
  const borderIndices = [
    Excel.BorderIndex.edgeTop,
    Excel.BorderIndex.edgeBottom,
    Excel.BorderIndex.edgeLeft,
    Excel.BorderIndex.edgeRight,
    Excel.BorderIndex.insideHorizontal,
    Excel.BorderIndex.insideVertical
  ];
  
  for (const index of borderIndices) {
    range.format.borders.getItem(index).style = Excel.BorderLineStyle.none;
  }
  
  await context.sync();
});
```

### Set Conditional Format
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.colorScale
  );
  
  conditionalFormat.colorScale.criteria = {
    minimum: { color: "#FF0000", type: "LowestValue" },
    midpoint: { color: "#FFFF00", type: "Percentile", value: 50 },
    maximum: { color: "#00FF00", type: "HighestValue" }
  };
  
  await context.sync();
});
```

## Formula Templates

### Common Formulas
```javascript
// Sum
"=SUM(A1:A10)"

// Average
"=AVERAGE(A1:A10)"

// Count
"=COUNT(A1:A10)"
"=COUNTA(A1:A10)"  // Non-empty count

// Conditional count
"=COUNTIF(A1:A10, \">100\")"

// Lookup
"=VLOOKUP(E1, A1:B10, 2, FALSE)"

// Conditional sum
"=SUMIF(A1:A10, \">100\", B1:B10)"
```

## Complete Conditional Formatting Templates

### Cell Value Conditional Format
```javascript
// Highlight cells greater than 100
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
  );
  
  // Set cells greater than 100 to green background
  conditionalFormat.cellValue.format.fill.color = "#90EE90";
  conditionalFormat.cellValue.rule = { 
    formula1: "100", 
    operator: "GreaterThan" 
  };
  
  await context.sync();
});
```

### Data Bar Format
```javascript
// Add data bars to range
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.dataBar
  );
  
  // Set data bar direction and color
  conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
  conditionalFormat.dataBar.positiveFormat.fillColor = "#4472C4";
  
  await context.sync();
});
```

### Icon Set Format
```javascript
// Display red-green arrow icon set
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("C2:C10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.iconSet
  );
  
  const iconSetCF = conditionalFormat.iconSet;
  iconSetCF.style = Excel.IconSet.threeArrows;
  
  // Set icon criteria
  iconSetCF.criteria = [
    {},
    {
      type: Excel.ConditionalFormatIconRuleType.number,
      operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
      formula: "=500"
    },
    {
      type: Excel.ConditionalFormatIconRuleType.number,
      operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual,
      formula: "=1000"
    }
  ];
  
  await context.sync();
});
```

### Preset Conditional Format (Above Average)
```javascript
// Highlight cells above average
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("D2:D10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.presetCriteria
  );
  
  // Set cells above average to yellow background
  conditionalFormat.preset.format.fill.color = "yellow";
  conditionalFormat.preset.rule = {
    criterion: Excel.ConditionalFormatPresetCriterion.aboveAverage
  };
  
  await context.sync();
});
```

### Top/Bottom N Format
```javascript
// Highlight top 10
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("E2:E20");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.topBottom
  );
  
  // Highlight top 10 items
  conditionalFormat.topBottom.format.fill.color = "#FFC000";
  conditionalFormat.topBottom.rule = {
    rank: 10,
    type: "TopItems"
  };
  
  await context.sync();
});
```

### Custom Formula Conditional Format
```javascript
// Set conditional format using custom formula
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.custom
  );
  
  // If cell value is greater than left cell, set to green
  conditionalFormat.custom.rule.formula = '=B2>A2';
  conditionalFormat.custom.format.font.color = "green";
  
  await context.sync();
});
```

## Event Handling Templates

### Monitor Cell Data Changes
```javascript
// Register data change event handler
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.onChanged.add(async (event) => {
    await Excel.run(async (context) => {
      console.log("Data change type:", event.changeType);
      console.log("Change address:", event.address);
      console.log("Change source:", event.source);
      await context.sync();
    });
  });
  
  await context.sync();
  console.log("Data change event listener registered");
});
```

### Monitor Selection Changes
```javascript
// Monitor user-selected cell changes
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.onSelectionChanged.add(async (event) => {
    await Excel.run(async (context) => {
      console.log("Current selection address:", event.address);
      await context.sync();
    });
  });
  
  await context.sync();
  console.log("Selection change event listener registered");
});
```

### Worksheet Activation Event
```javascript
// Monitor worksheet activation
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  
  sheets.onActivated.add(async (event) => {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(event.worksheetId);
      sheet.load("name");
      await context.sync();
      console.log("Activated worksheet:", sheet.name);
    });
  });
  
  await context.sync();
  console.log("Worksheet activation event listener registered");
});
```

### Calculation Complete Event
```javascript
// Monitor worksheet calculation completion
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.onCalculated.add(async (event) => {
    await Excel.run(async (context) => {
      console.log("Worksheet calculation complete:", event.worksheetId);
      await context.sync();
    });
  });
  
  await context.sync();
  console.log("Calculation complete event listener registered");
});
```

## Shape Operation Templates

### Add Rectangle Shape
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  const rectangle = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
  rectangle.left = 100;
  rectangle.top = 100;
  rectangle.height = 150;
  rectangle.width = 200;
  rectangle.name = "MyRectangle";
  
  // Set fill color
  rectangle.fill.setSolidColor("#4472C4");
  
  await context.sync();
});
```

### Insert Image
```javascript
// Insert Base64 encoded image
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // Sample Base64 image (replace with actual image data)
  const base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
  
  const image = shapes.addImage(base64Image);
  image.left = 200;
  image.top = 50;
  image.height = 100;
  image.width = 100;
  image.name = "MyImage";
  
  await context.sync();
});
```

### Add Line
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // Add straight line (from point [200,50] to [300,150])
  const line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
  line.name = "MyLine";
  
  // Set line style
  line.lineFormat.color = "red";
  line.lineFormat.weight = 2;
  
  await context.sync();
});
```

### Create Text Box
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  const textBox = shapes.addTextBox("This is text box content");
  textBox.left = 100;
  textBox.top = 200;
  textBox.height = 80;
  textBox.width = 200;
  textBox.name = "MyTextBox";
  
  // Set text format
  textBox.textFrame.textRange.font.color = "blue";
  textBox.textFrame.textRange.font.size = 14;
  
  await context.sync();
});
```

### Group Shapes
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // Get shapes to group
  const shape1 = shapes.getItem("Shape1");
  const shape2 = shapes.getItem("Shape2");
  
  // Create shape group
  shape1.load("id");
  shape2.load("id");
  await context.sync();
  
  const shapeIds = [shape1.id, shape2.id];
  const group = shapes.addGroup(shapeIds);
  group.name = "MyGroup";
  
  await context.sync();
});
```

## Notes Templates

### Add Cell Note
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Add note at A1 cell
  sheet.notes.add("A1", "This is an important note.");
  
  await context.sync();
});
```

### Modify Note Content
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Get and modify note at A1
  const note = sheet.notes.getItem("A1");
  note.content = "Updated note content.";
  
  await context.sync();
});
```

### Set Note Visibility
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  const note = sheet.notes.getItem("A1");
  note.load("visible");
  await context.sync();
  
  // Toggle visibility
  note.visible = !note.visible;
  
  await context.sync();
});
```

### Delete Note
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  const note = sheet.notes.getItem("A2");
  note.delete();
  
  await context.sync();
});
```

## Range Advanced Operation Templates

### Copy Range to New Location
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Copy A1:C5 to E1
  const destRange = sheet.getRange("E1");
  destRange.copyFrom("A1:C5", Excel.RangeCopyType.all);
  
  await context.sync();
});
```

### Copy with Skip Blanks
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Skip blank cells when copying, preserve existing data at destination
  const destRange = sheet.getRange("D1");
  destRange.copyFrom(
    "A1:C3",
    Excel.RangeCopyType.all,
    true,  // skipBlanks
    false  // transpose
  );
  
  await context.sync();
});
```

### Move Range
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Move A1:C5 to G1 (cut and paste)
  const sourceRange = sheet.getRange("A1:C5");
  sourceRange.moveTo("G1");
  
  await context.sync();
});
```

### Insert Cells
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Insert cell at B2, existing cells shift down
  const range = sheet.getRange("B2:B2");
  range.insert(Excel.InsertShiftDirection.down);
  
  await context.sync();
});
```

### Delete Cells
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Delete C3:C5, other cells shift up
  const range = sheet.getRange("C3:C5");
  range.delete(Excel.DeleteShiftDirection.up);
  
  await context.sync();
});
```

### Clear Range Content
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // Clear all content and formats
  range.clear(Excel.ClearApplyTo.all);
  
  await context.sync();
});
```

### Remove Duplicates
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:B20");
  
  // Remove duplicate rows based on columns 1 and 2
  range.removeDuplicates([0, 1], true); // true means includes header row
  
  await context.sync();
});
```

### Group Rows/Columns
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Group rows 2 to 5
  const range = sheet.getRange("2:5");
  range.group(Excel.GroupOption.byRows);
  
  await context.sync();
});
```

### Ungroup
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // Ungroup rows 2 to 5
  const range = sheet.getRange("2:5");
  range.ungroup(Excel.GroupOption.byRows);
  
  await context.sync();
});
```

## Checkbox Templates

### Add Checkbox
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // Convert boolean values to checkboxes
  range.control = {
    type: Excel.CellControlType.checkbox
  };
  
  await context.sync();
});
```

### Read Checkbox State
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // Read checkbox values (true/false)
  range.load("values");
  await context.sync();
  
  console.log("Checkbox state:", range.values);
});
```

### Remove Checkbox
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // Convert checkbox back to boolean value
  range.control = {
    type: Excel.CellControlType.empty
  };
  
  await context.sync();
});
```

## Data Analysis Python Templates

### Descriptive Statistical Analysis
```python
import pandas as pd
import numpy as np

def analyze_data(data):
    """Perform descriptive statistical analysis on Excel data"""
    df = pd.DataFrame(data[1:], columns=data[0])
    
    stats = {
        "count": len(df),
        "mean": df.select_dtypes(include=[np.number]).mean().to_dict(),
        "std": df.select_dtypes(include=[np.number]).std().to_dict(),
        "min": df.select_dtypes(include=[np.number]).min().to_dict(),
        "max": df.select_dtypes(include=[np.number]).max().to_dict(),
        "median": df.select_dtypes(include=[np.number]).median().to_dict(),
    }
    
    return stats
```

### Data Cleaning
```python
def clean_data(data):
    """Clean Excel data"""
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # Remove duplicate rows
    df = df.drop_duplicates()
    
    # Fill missing values
    df = df.fillna(method='ffill')
    
    # Remove whitespace
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()
    
    return df.values.tolist()
```
