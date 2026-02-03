/**
 * Word Table Actions 代码生成器
 */

/**
 * 生成创建表格代码
 */
export function generateCreateCode(params: {
  rows: number;
  columns: number;
  data?: any[][];
  location?: string;
  style?: string;
}): string {
  const location = params.location || 'End';
  const dataJson = params.data ? JSON.stringify(params.data) : 'null';
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  
  // 创建表格
  const table = body.insertTable(${params.rows}, ${params.columns}, Word.InsertLocation.${location.toLowerCase()});
  
  ${params.data ? `
  // 填充数据
  const tableData = ${dataJson};
  for (let i = 0; i < tableData.length && i < ${params.rows}; i++) {
    for (let j = 0; j < tableData[i].length && j < ${params.columns}; j++) {
      table.getCell(i, j).body.insertText(String(tableData[i][j]), Word.InsertLocation.replace);
    }
  }
  ` : ''}
  
  ${params.style ? `table.style = "${params.style}";` : ''}
  
  table.load("values");
  await context.sync();
  
  // 获取表格索引
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  let tableIndex = -1;
  for (let i = 0; i < tables.items.length; i++) {
    if (tables.items[i] === table) {
      tableIndex = i;
      break;
    }
  }
  
  return {
    tableIndex: tableIndex,
    success: true
  };
});`.trim();
}

/**
 * 生成读取表格代码
 */
export function generateReadCode(params: { tableIndex: number }): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  table.load("rowCount, columnCount, values");
  await context.sync();
  
  return {
    data: table.values,
    rows: table.rowCount,
    columns: table.columnCount
  };
});`.trim();
}

/**
 * 生成写入表格数据代码
 */
export function generateWriteCode(params: {
  tableIndex: number;
  data: any[][];
}): string {
  const dataJson = JSON.stringify(params.data);
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  const tableData = ${dataJson};
  
  for (let i = 0; i < tableData.length; i++) {
    for (let j = 0; j < tableData[i].length; j++) {
      table.getCell(i, j).body.insertText(String(tableData[i][j]), Word.InsertLocation.replace);
    }
  }
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成插入行代码
 */
export function generateInsertRowCode(params: {
  tableIndex: number;
  rowIndex: number;
  values?: any[];
}): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  const row = table.insertRows(Word.InsertLocation.${params.rowIndex === 0 ? 'start' : 'after'}, 1, ${params.values ? JSON.stringify([params.values]) : 'null'});
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成插入列代码
 */
export function generateInsertColumnCode(params: {
  tableIndex: number;
  columnIndex: number;
  values?: any[];
}): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  table.addColumns(Word.InsertLocation.${params.columnIndex === 0 ? 'start' : 'after'}, 1, ${params.values ? JSON.stringify(params.values) : 'null'});
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除行代码
 */
export function generateDeleteRowCode(params: {
  tableIndex: number;
  rowIndex: number;
}): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  const row = table.rows.items[${params.rowIndex}];
  row.delete();
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除列代码
 */
export function generateDeleteColumnCode(params: {
  tableIndex: number;
  columnIndex: number;
}): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  table.deleteColumns(${params.columnIndex}, 1);
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成格式化单元格代码
 */
export function generateFormatCellCode(params: {
  tableIndex: number;
  rowIndex: number;
  columnIndex: number;
  format: any;
}): string {
  const formatters: string[] = [];
  
  if (params.format.fill?.color) {
    formatters.push(`cell.shadingColor = "${params.format.fill.color}";`);
  }
  
  if (params.format.font) {
    const font = params.format.font;
    if (font.size) formatters.push(`cell.body.font.size = ${font.size};`);
    if (font.bold !== undefined) formatters.push(`cell.body.font.bold = ${font.bold};`);
    if (font.color) formatters.push(`cell.body.font.color = "${font.color}";`);
  }
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  const table = tables.items[${params.tableIndex}];
  const cell = table.getCell(${params.rowIndex}, ${params.columnIndex});
  
  ${formatters.join('\n  ')}
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除表格代码
 */
export function generateDeleteTableCode(params: { tableIndex: number }): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const tables = body.tables;
  tables.load("items");
  await context.sync();
  
  if (${params.tableIndex} < 0 || ${params.tableIndex} >= tables.items.length) {
    return {
      success: false,
      error: "表格索引越界: ${params.tableIndex}"
    };
  }
  
  tables.items[${params.tableIndex}].delete();
  await context.sync();
  
  return {
    success: true,
    deleted: ${params.tableIndex}
  };
});`.trim();
}
