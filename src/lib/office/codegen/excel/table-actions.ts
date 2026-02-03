/**
 * Excel Table Actions 代码生成器
 */

/**
 * 生成创建表格代码
 */
export function generateCreateCode(params: {
  address: string;
  hasHeaders: boolean;
  name?: string;
  style?: string;
}): string {
  const tableName = params.name || `Table${Date.now()}`;
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.address}");
  
  // 创建表格
  const table = sheet.tables.add(range, ${params.hasHeaders});
  table.name = "${tableName}";
  ${params.style ? `table.style = "${params.style}";` : ''}
  
  table.load("name");
  await context.sync();
  
  return {
    tableName: table.name,
    success: true
  };
});`.trim();
}

/**
 * 生成读取表格代码
 */
export function generateReadCode(params: {
  name: string;
  includeHeaders?: boolean;
}): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  const dataRange = table.getDataBodyRange();
  dataRange.load("values");
  
  ${params.includeHeaders !== false ? `
  const headerRange = table.getHeaderRowRange();
  headerRange.load("values");
  ` : ''}
  
  await context.sync();
  
  return {
    name: "${params.name}",
    ${params.includeHeaders !== false ? 'headers: headerRange.values[0],' : ''}
    data: dataRange.values,
    totalRows: dataRange.values.length
  };
});`.trim();
}

/**
 * 生成添加行代码
 */
export function generateAddRowCode(params: {
  name: string;
  values?: any[];
  index?: number;
}): string {
  const valuesJson = params.values ? JSON.stringify([params.values]) : 'null';
  const indexParam = params.index !== undefined ? params.index : 'null';
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  const newRow = table.rows.add(${indexParam}, ${valuesJson});
  newRow.load("index");
  await context.sync();
  
  return {
    success: true,
    newRowIndex: newRow.index
  };
});`.trim();
}

/**
 * 生成添加列代码
 */
export function generateAddColumnCode(params: {
  name: string;
  columnName: string;
  values?: any[];
}): string {
  const valuesJson = params.values ? JSON.stringify(params.values) : 'null';
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  const newColumn = table.columns.add(null, ${valuesJson}, "${params.columnName}");
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成排序代码
 */
export function generateSortCode(params: {
  name: string;
  column: string;
  ascending: boolean;
}): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  const column = table.columns.getItemOrNullObject("${params.column}");
  await context.sync();
  
  if (column.isNullObject) {
    return {
      success: false,
      error: "列不存在: ${params.column}"
    };
  }
  
  table.sort.apply([{
    key: column.index,
    ascending: ${params.ascending}
  }]);
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成筛选代码
 */
export function generateFilterCode(params: {
  name: string;
  column: string;
  criteria: any;
}): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  const column = table.columns.getItemOrNullObject("${params.column}");
  await context.sync();
  
  if (column.isNullObject) {
    return {
      success: false,
      error: "列不存在: ${params.column}"
    };
  }
  
  // 应用筛选（简化版本）
  table.autoFilter.apply(table.getRange());
  
  await context.sync();
  
  return {
    success: true,
    note: "已启用自动筛选"
  };
});`.trim();
}

/**
 * 生成删除表格代码
 */
export function generateDeleteCode(params: { name: string }): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const table = sheet.tables.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (table.isNullObject) {
    return {
      success: false,
      error: "表格不存在: ${params.name}"
    };
  }
  
  table.delete();
  await context.sync();
  
  return {
    success: true,
    deleted: "${params.name}"
  };
});`.trim();
}
