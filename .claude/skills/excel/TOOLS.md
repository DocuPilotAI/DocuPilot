# Excel 工具模板库

## ⚠️ 重要：优先使用 MCP 领域工具

**本文件仅包含供参考的低层 Office.js 代码模板。**

**实际开发中请优先使用 MCP 领域工具：**
- `excel_range` - 单元格读写、格式设置操作
- `excel_worksheet` - 工作表管理（添加、删除、重命名、激活）
- `excel_table` - 表格对象操作（创建、排序、筛选）
- `excel_chart` - 图表创建与管理

**仅对以下情况使用 execute_code + 本文件模板：**
- 数据透视表操作
- 条件格式（复杂规则）
- 数据验证（高级验证逻辑）
- 事件处理（工作表事件）
- MCP 工具未覆盖的其他高级 API

**性能对比：**
- MCP 工具：1.2s 响应，~280 tokens，<5% 错误率
- execute_code：2.5s 响应，~800 tokens，15% 错误率

**另见：**
- [MCP 工具 API 文档](../../../docs/MCP_TOOLS_API.md)
- [MCP 工具决策流程](../../../docs/MCP_TOOL_DECISION_FLOW.md)

---

## 工作表管理模板

### 创建工作表
```javascript
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.add("NewSheet");
  sheet.activate();
  await context.sync();
});
```

### 创建工作表（带验证）
```javascript
// ⚠️ 最佳实践：检查工作表是否已存在
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const existingSheet = sheets.getItemOrNullObject("NewSheet");
  await context.sync();
  
  let sheet;
  if (existingSheet.isNullObject) {
    // 工作表不存在，创建新的
    sheet = sheets.add("NewSheet");
  } else {
    // 工作表已存在，使用现有
    sheet = existingSheet;
  }
  
  sheet.activate();
  await context.sync();
});
```

### 重命名工作表
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.name = "RenamedSheet";
  await context.sync();
});
```

### 删除工作表
```javascript
// ⚠️ 重要：删除前检查工作表是否存在，建议使用「删除工作表（带验证）」模板
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

### 删除工作表（带验证）
```javascript
// ⚠️ 最佳实践：删除前检查工作表存在性（推荐）
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItemOrNullObject("SheetToDelete");
  await context.sync();
  
  if (!sheet.isNullObject) {
    sheet.delete();
    await context.sync();
    console.log("工作表已删除");
  } else {
    console.log("工作表不存在，无需删除");
  }
});
```

### 复制工作表
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const newSheet = sheet.copy(Excel.WorksheetPositionType.after, sheet);
  newSheet.name = "CopiedSheet";
  await context.sync();
});
```

## 数据读取模板

### 读取选中区域
```javascript
// 获取当前选中的单元格区域
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

### 读取指定区域
```javascript
// 读取指定地址的单元格
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  range.load("values");
  await context.sync();
  
  return range.values;
});
```

### 读取多个非连续区域（重要！）
```javascript
// ⚠️ 注意：getRangeAreas 是工作簿级方法，非工作表级
// 错误用法：sheet.getRangeAreas("B3,F3,J3") ❌
// 正确用法：workbook.getRangeAreas() ✅

Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sheetName = sheet.name;
  
  // 使用 workbook.getRangeAreas() 读取多个非连续区域
  const rangeAreas = context.workbook.getRangeAreas(`${sheetName}!B3,${sheetName}!F3,${sheetName}!J3`);
  rangeAreas.load("address");
  
  // 或分别读取每个单元格
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

## 数据写入模板

### 写入单个值
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const cell = sheet.getRange("A1");
  cell.values = [["Hello World"]];
  await context.sync();
});
```

### 写入数组数据
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

### 写入多个非连续区域
```javascript
// 方法 1：分别写入每个区域（推荐）
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  sheet.getRange("B3").values = [[10]];
  sheet.getRange("F3").values = [[20]];
  sheet.getRange("J3").values = [[30]];
  
  await context.sync();
});

// 方法 2：使用 workbook.getRangeAreas()（跨工作表场景）
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sheetName = sheet.name;
  
  // 注意：必须使用完整地址（含工作表名）
  const rangeAreas = context.workbook.getRangeAreas(
    `${sheetName}!B3,${sheetName}!F3,${sheetName}!J3`
  );
  
  // RangeAreas 操作受限，通常用于格式设置和批量操作
  rangeAreas.format.fill.color = "yellow";
  
  await context.sync();
});
```

## 表格操作模板

### 创建表格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  // 创建带表头的表格
  const table = sheet.tables.add("A1:D5", true);
  table.name = "SalesTable";
  
  // 设置表头
  table.getHeaderRowRange().values = [["Date", "Product", "Category", "Amount"]];
  
  // Set data
  table.getDataBodyRange().values = [
    ["2023-01-01", "Widget A", "Electronics", 100],
    ["2023-01-02", "Widget B", "Electronics", 200],
    ["2023-01-03", "Widget C", "Home", 150],
    ["2023-01-04", "Widget D", "Home", 300]
  ];
  
  // 自动调整列宽
  sheet.getUsedRange().format.autofitColumns();
  
  await context.sync();
});
```

### 排序表格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItem("SalesTable");
  
  // 按第 4 列（Amount）降序排序
  table.sort.apply([
    {
      key: 3, // 第 4 列，索引从 0 开始
      ascending: false
    }
  ]);
  
  await context.sync();
});
```

### 筛选表格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItem("SalesTable");
  
  // 筛选第 3 列（Category）为 "Electronics"
  table.columns.getItemAt(2).filter.apply({
    filterOn: Excel.FilterOn.values,
    values: ["Electronics"]
  });
  
  await context.sync();
});
```

## 数据透视表模板

### 创建数据透视表
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  // 假设数据在 A1:D5
  const sourceRange = sheet.getRange("A1:D5");
  
  // 在 F1 创建数据透视表
  const pivotTable = sheet.pivotTables.add("PivotTable1", sourceRange, "F1");
  
  // 添加行字段
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Category"));
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Product"));
  
  // 添加数据字段（默认求和）
  pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Amount"));
  
  await context.sync();
});
```

### 添加列层级和筛选
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 添加列字段
  pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Region"));
  
  // 添加筛选字段
  pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem("Category"));
  
  await context.sync();
});
```

### 数据透视表日期筛选
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 获取或添加日期层级
  let dateHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Date");
  await context.sync();
  
  if (dateHierarchy.isNullObject) {
    dateHierarchy = pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Date"));
    await context.sync();
  }
  
  // 应用日期筛选：仅显示 2020-08-01 之后的数据
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

### 数据透视表标签筛选
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 获取字段
  const filterField = pivotTable.rowHierarchies.getItem("Category")
    .fields.getItem("Category");
  
  // 标签筛选：排除以 "Electronics" 开头的项
  const labelFilter = {
    condition: Excel.LabelFilterCondition.beginsWith,
    substring: "Electronics",
    exclusive: true
  };
  filterField.applyFilter({ labelFilter: labelFilter });
  
  await context.sync();
});
```

### 数据透视表值筛选
```javascript
// ⚠️ 值筛选：按聚合值筛选行字段项
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
  await context.sync();
  
  if (pivotTable.isNullObject) {
    console.log("数据透视表不存在");
    return;
  }
  
  // 获取行层级中的字段
  const productHierarchy = pivotTable.rowHierarchies.getItemOrNullObject("Product");
  await context.sync();
  
  if (!productHierarchy.isNullObject) {
    const filterField = productHierarchy.fields.getItem("Product");
    
    // 值筛选：仅显示销售额大于 500 的项
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

### 清除数据透视表筛选
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 加载所有层级
  pivotTable.hierarchies.load("items");
  await context.sync();
  
  // 清除所有字段的筛选
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

### 创建切片器
```javascript
// ⚠️ 创建用于交互式数据透视表筛选的切片器
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 首先检查数据透视表是否存在
  const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
  await context.sync();
  
  if (pivotTable.isNullObject) {
    console.log("数据透视表不存在，无法创建切片器");
    return;
  }
  
  // 创建用于交互式筛选的切片器
  const slicer = sheet.slicers.add(
    "PivotTable1",  // 数据透视表名称
    "Category"      // 筛选字段
  );
  
  slicer.name = "Category Slicer";
  slicer.left = 400;
  slicer.top = 200;
  slicer.width = 200;
  slicer.height = 200;
  
  await context.sync();
});
```

### 使用切片器筛选
```javascript
Excel.run(async (context) => {
  const slicer = context.workbook.slicers.getItem("Category Slicer");
  
  // 选择要筛选的特定项
  slicer.selectItems(["Electronics", "Home", "Clothing"]);
  
  await context.sync();
});
```

### 切换数据透视表布局
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 切换布局类型：紧凑、大纲、表格
  pivotTable.layout.layoutType = "Outline";  // or "Compact", "Tabular"
  
  await context.sync();
});
```

### 获取数据透视表数据
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 获取数据区域
  const dataRange = pivotTable.layout.getDataBodyRange();
  dataRange.load("address, values");
  
  await context.sync();
  
  console.log("Data range:", dataRange.address);
  console.log("Data values:", dataRange.values);
  
  return dataRange.values;
});
```

### 格式化数据透视表
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  const pivotLayout = pivotTable.layout;
  
  // 设置空单元格显示文本
  pivotLayout.emptyCellText = "--";
  pivotLayout.fillEmptyCells = true;
  
  // 保留格式设置
  pivotLayout.preserveFormatting = true;
  
  // 设置数据区域对齐
  const dataRange = pivotLayout.getDataBodyRange();
  dataRange.format.horizontalAlignment = Excel.HorizontalAlignment.right;
  
  await context.sync();
});
```

### 刷新数据透视表
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 刷新数据透视表数据
  pivotTable.refresh();
  
  await context.sync();
});
```

### 删除数据透视表
```javascript
Excel.run(async (context) => {
  const pivotTable = context.workbook.worksheets.getActiveWorksheet()
    .pivotTables.getItem("PivotTable1");
  
  // 删除数据透视表
  pivotTable.delete();
  
  await context.sync();
});
```

## 数据验证模板

### 设置数值验证
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // 仅允许大于 0 的整数
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: 0,
      operator: Excel.DataValidationOperator.greaterThan
    }
  };
  
  // 设置输入提示
  range.dataValidation.prompt = {
    showPrompt: true,
    title: "输入限制",
    message: "请输入大于 0 的整数"
  };
  
  // 设置错误提示
  range.dataValidation.errorAlert = {
    showAlert: true,
    style: Excel.DataValidationAlertStyle.stop,
    title: "输入无效",
    message: "值必须大于 0"
  };
  
  await context.sync();
});
```

### 设置下拉列表
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

## 评论与命名项

### 添加评论
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const comments = sheet.comments;
  
  // 在 A1 添加评论
  comments.add("A1", "This is a comment.");
  
  await context.sync();
});
```

### 创建命名区域
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:B10");
  
  // 创建命名区域 "MyData"
  context.workbook.names.add("MyData", range);
  
  await context.sync();
});
```

## 图表创建模板

### 柱状图
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

### 折线图
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

## 格式设置模板

### 字体格式
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // 字体属性在 format.font 下
  range.format.font.name = "Arial";
  range.format.font.size = 12;
  range.format.font.bold = true;
  range.format.font.italic = false;
  range.format.font.underline = Excel.RangeUnderlineStyle.single;
  range.format.font.color = "#FF0000";
  
  await context.sync();
});
```

### 填充颜色
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // 填充颜色在 format.fill 下
  range.format.fill.color = "#FFFF00"; // 黄色背景
  
  await context.sync();
});
```

### 数字格式
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // 数字格式使用二维数组结构
  range.numberFormat = [["#,##0.00"]]; // 千分位分隔符，两位小数
  // 其他示例：
  // range.numberFormat = [["0.00%"]]; // 百分比
  // range.numberFormat = [["$#,##0.00"]]; // 货币
  // range.numberFormat = [["m/d/yyyy"]]; // 日期格式
  
  await context.sync();
});
```

### 对齐
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // 对齐属性直接在 format 下
  range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
  range.format.verticalAlignment = Excel.VerticalAlignment.center;
  range.format.wrapText = true;
  range.format.textOrientation = 0; // 0-90 度
  
  await context.sync();
});
```

### 边框格式
```javascript
// ⚠️ API 规则：使用 format.borders（复数集合），而非 format.border（不存在）
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // 可用的 BorderIndex 值：
  // - edgeTop, edgeBottom, edgeLeft, edgeRight（外边框）
  // - insideHorizontal, insideVertical（内边框）
  // - diagonalDown, diagonalUp（对角线）
  
  // 设置外边框
  range.format.borders.getItem(Excel.BorderIndex.edgeTop).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeBottom).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeLeft).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.edgeRight).style = Excel.BorderLineStyle.continuous;
  
  // 设置内边框（仅适用于多单元格区域）
  range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).style = Excel.BorderLineStyle.continuous;
  range.format.borders.getItem(Excel.BorderIndex.insideVertical).style = Excel.BorderLineStyle.continuous;
  
  // 自定义边框属性
  const topBorder = range.format.borders.getItem(Excel.BorderIndex.edgeTop);
  topBorder.color = "#000000";
  topBorder.weight = Excel.BorderWeight.thick; // thin, medium, thick, hairline
  
  await context.sync();
});
```

### 边框格式（所有边框循环）
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // 使用循环对所有边框应用相同样式
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

### 清除边框
```javascript
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  
  // 通过将样式设为 none 清除所有边框
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

### 设置条件格式
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

## 公式模板

### 常用公式
```javascript
// 求和
"=SUM(A1:A10)"

// 平均值
"=AVERAGE(A1:A10)"

// 计数
"=COUNT(A1:A10)"
"=COUNTA(A1:A10)"  // 非空计数

// 条件计数
"=COUNTIF(A1:A10, \">100\")"

// 查找
"=VLOOKUP(E1, A1:B10, 2, FALSE)"

// 条件求和
"=SUMIF(A1:A10, \">100\", B1:B10)"
```

## 完整条件格式模板

### 单元格值条件格式
```javascript
// 高亮大于 100 的单元格
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.cellValue
  );
  
  // 将大于 100 的单元格设置为绿色背景
  conditionalFormat.cellValue.format.fill.color = "#90EE90";
  conditionalFormat.cellValue.rule = { 
    formula1: "100", 
    operator: "GreaterThan" 
  };
  
  await context.sync();
});
```

### 数据条格式
```javascript
// 向区域添加数据条
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.dataBar
  );
  
  // 设置数据条方向和颜色
  conditionalFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
  conditionalFormat.dataBar.positiveFormat.fillColor = "#4472C4";
  
  await context.sync();
});
```

### 图标集格式
```javascript
// 显示红绿箭头图标集
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("C2:C10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.iconSet
  );
  
  const iconSetCF = conditionalFormat.iconSet;
  iconSetCF.style = Excel.IconSet.threeArrows;
  
  // 设置图标条件
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

### 预设条件格式（高于平均值）
```javascript
// 高亮高于平均值的单元格
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("D2:D10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.presetCriteria
  );
  
  // 将高于平均值的单元格设置为黄色背景
  conditionalFormat.preset.format.fill.color = "yellow";
  conditionalFormat.preset.rule = {
    criterion: Excel.ConditionalFormatPresetCriterion.aboveAverage
  };
  
  await context.sync();
});
```

### 前 N/后 N 格式
```javascript
// 高亮前 10
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("E2:E20");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.topBottom
  );
  
  // 高亮前 10 项
  conditionalFormat.topBottom.format.fill.color = "#FFC000";
  conditionalFormat.topBottom.rule = {
    rank: 10,
    type: "TopItems"
  };
  
  await context.sync();
});
```

### 自定义公式条件格式
```javascript
// 使用自定义公式设置条件格式
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  const conditionalFormat = range.conditionalFormats.add(
    Excel.ConditionalFormatType.custom
  );
  
  // 若单元格值大于左侧单元格，设为绿色
  conditionalFormat.custom.rule.formula = '=B2>A2';
  conditionalFormat.custom.format.font.color = "green";
  
  await context.sync();
});
```

## 事件处理模板

### 监控单元格数据变化
```javascript
// 注册数据变化事件处理器
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

### 监控选择变化
```javascript
// 监控用户选中的单元格变化
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

### 工作表激活事件
```javascript
// 监控工作表激活
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

### 计算完成事件
```javascript
// 监控工作表计算完成
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

## 形状操作模板

### 添加矩形形状
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
  
  // 设置填充颜色
  rectangle.fill.setSolidColor("#4472C4");
  
  await context.sync();
});
```

### 插入图片
```javascript
// 插入 Base64 编码的图片
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // 示例 Base64 图片（替换为实际图片数据）
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

### 添加线条
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // 添加直线（从点 [200,50] 到 [300,150]）
  const line = shapes.addLine(200, 50, 300, 150, Excel.ConnectorType.straight);
  line.name = "MyLine";
  
  // 设置线条样式
  line.lineFormat.color = "red";
  line.lineFormat.weight = 2;
  
  await context.sync();
});
```

### 创建文本框
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
  
  // 设置文本格式
  textBox.textFrame.textRange.font.color = "blue";
  textBox.textFrame.textRange.font.size = 14;
  
  await context.sync();
});
```

### 组合形状
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const shapes = sheet.shapes;
  
  // 获取要组合的形状
  const shape1 = shapes.getItem("Shape1");
  const shape2 = shapes.getItem("Shape2");
  
  // 创建形状组
  shape1.load("id");
  shape2.load("id");
  await context.sync();
  
  const shapeIds = [shape1.id, shape2.id];
  const group = shapes.addGroup(shapeIds);
  group.name = "MyGroup";
  
  await context.sync();
});
```

## 批注模板

### 添加单元格批注
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 在 A1 单元格添加批注
  sheet.notes.add("A1", "This is an important note.");
  
  await context.sync();
});
```

### 修改批注内容
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 获取并修改 A1 的批注
  const note = sheet.notes.getItem("A1");
  note.content = "Updated note content.";
  
  await context.sync();
});
```

### 设置批注可见性
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  const note = sheet.notes.getItem("A1");
  note.load("visible");
  await context.sync();
  
  // 切换可见性
  note.visible = !note.visible;
  
  await context.sync();
});
```

### 删除批注
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  const note = sheet.notes.getItem("A2");
  note.delete();
  
  await context.sync();
});
```

## 区域高级操作模板

### 复制区域到新位置
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 将 A1:C5 复制到 E1
  const destRange = sheet.getRange("E1");
  destRange.copyFrom("A1:C5", Excel.RangeCopyType.all);
  
  await context.sync();
});
```

### 复制时跳过空单元格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 复制时跳过空单元格，保留目标处的现有数据
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

### 移动区域
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 将 A1:C5 移动到 G1（剪切粘贴）
  const sourceRange = sheet.getRange("A1:C5");
  sourceRange.moveTo("G1");
  
  await context.sync();
});
```

### 插入单元格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 在 B2 插入单元格，现有单元格下移
  const range = sheet.getRange("B2:B2");
  range.insert(Excel.InsertShiftDirection.down);
  
  await context.sync();
});
```

### 删除单元格
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 删除 C3:C5，其他单元格上移
  const range = sheet.getRange("C3:C5");
  range.delete(Excel.DeleteShiftDirection.up);
  
  await context.sync();
});
```

### 清除区域内容
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:D10");
  
  // 清除所有内容和格式
  range.clear(Excel.ClearApplyTo.all);
  
  await context.sync();
});
```

### 删除重复项
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:B20");
  
  // 根据第 1 列和第 2 列删除重复行
  range.removeDuplicates([0, 1], true); // true 表示包含表头行
  
  await context.sync();
});
```

### 行列分组
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 对第 2 到 5 行分组
  const range = sheet.getRange("2:5");
  range.group(Excel.GroupOption.byRows);
  
  await context.sync();
});
```

### 取消分组
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  
  // 取消第 2 到 5 行的分组
  const range = sheet.getRange("2:5");
  range.ungroup(Excel.GroupOption.byRows);
  
  await context.sync();
});
```

## 复选框模板

### 添加复选框
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // 将布尔值转换为复选框
  range.control = {
    type: Excel.CellControlType.checkbox
  };
  
  await context.sync();
});
```

### 读取复选框状态
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // 读取复选框值（true/false）
  range.load("values");
  await context.sync();
  
  console.log("Checkbox state:", range.values);
});
```

### 移除复选框
```javascript
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B2:B10");
  
  // 将复选框恢复为布尔值
  range.control = {
    type: Excel.CellControlType.empty
  };
  
  await context.sync();
});
```

## 数据分析 Python 模板

### 描述性统计分析
```python
import pandas as pd
import numpy as np

def analyze_data(data):
    """对 Excel 数据执行描述性统计分析"""
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

### 数据清洗
```python
def clean_data(data):
    """清洗 Excel 数据"""
    df = pd.DataFrame(data[1:], columns=data[0])
    
    # 删除重复行
    df = df.drop_duplicates()
    
    # 填充缺失值
    df = df.fillna(method='ffill')
    
    # 去除空白
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].str.strip()
    
    return df.values.tolist()
```
