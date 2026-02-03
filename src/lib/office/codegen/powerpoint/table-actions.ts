/**
 * PowerPoint Table Actions 代码生成器
 */

/**
 * 生成创建表格代码
 */
export function generateCreateCode(params: {
  slideIndex: number;
  rows: number;
  columns: number;
  position: any;
  data?: any[][];
  style?: any;
}): string {
  const { left, top, width, height } = params.position;
  
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (${params.slideIndex} < 0 || ${params.slideIndex} >= slides.items.length) {
    return {
      success: false,
      error: "幻灯片索引越界: ${params.slideIndex}"
    };
  }
  
  const slide = slides.items[${params.slideIndex}];
  
  // 创建表格
  const table = slide.shapes.addTable(${params.rows}, ${params.columns}, {
    left: ${left},
    top: ${top},
    width: ${width || 600},
    height: ${height || 300}
    ${params.data ? `,\n    values: ${JSON.stringify(params.data)}` : ''}
  });
  
  table.load("id");
  await context.sync();
  
  return {
    tableId: table.id,
    success: true
  };
});`.trim();
}

/**
 * 生成读取表格代码
 */
export function generateReadCode(params: {
  slideIndex: number;
  tableId: string;
}): string {
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (${params.slideIndex} < 0 || ${params.slideIndex} >= slides.items.length) {
    return {
      success: false,
      error: "幻灯片索引越界: ${params.slideIndex}"
    };
  }
  
  const slide = slides.items[${params.slideIndex}];
  const shape = slide.shapes.getItem("${params.tableId}");
  
  // 检查是否为表格
  shape.load("type");
  await context.sync();
  
  if (shape.type !== PowerPoint.ShapeType.table) {
    return {
      success: false,
      error: "指定的形状不是表格"
    };
  }
  
  const table = shape.table;
  table.load("rows, columns");
  await context.sync();
  
  // 读取所有单元格数据
  const data = [];
  for (let i = 0; i < table.rows.items.length; i++) {
    const rowData = [];
    for (let j = 0; j < table.columns.items.length; j++) {
      const cell = table.getCell(i, j);
      cell.load("textFrame");
      await context.sync();
      rowData.push(cell.textFrame.textRange.text);
    }
    data.push(rowData);
  }
  
  return {
    data: data,
    rows: table.rows.items.length,
    columns: table.columns.items.length
  };
});`.trim();
}

/**
 * 生成写入表格数据代码
 */
export function generateWriteCode(params: {
  slideIndex: number;
  tableId: string;
  data: any[][];
}): string {
  const dataJson = JSON.stringify(params.data);
  
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (${params.slideIndex} < 0 || ${params.slideIndex} >= slides.items.length) {
    return {
      success: false,
      error: "幻灯片索引越界: ${params.slideIndex}"
    };
  }
  
  const slide = slides.items[${params.slideIndex}];
  const shape = slide.shapes.getItem("${params.tableId}");
  const table = shape.table;
  
  const tableData = ${dataJson};
  
  for (let i = 0; i < tableData.length; i++) {
    for (let j = 0; j < tableData[i].length; j++) {
      const cell = table.getCell(i, j);
      cell.textFrame.textRange.text = String(tableData[i][j]);
    }
  }
  
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
  slideIndex: number;
  tableId: string;
  row: number;
  column: number;
  format: any;
}): string {
  const formatters: string[] = [];
  
  if (params.format.fill?.color) {
    formatters.push(`cell.fill.setSolidColor("${params.format.fill.color}");`);
  }
  
  if (params.format.font) {
    const font = params.format.font;
    if (font.size) formatters.push(`cell.textFrame.textRange.font.size = ${font.size};`);
    if (font.bold !== undefined) formatters.push(`cell.textFrame.textRange.font.bold = ${font.bold};`);
    if (font.color) formatters.push(`cell.textFrame.textRange.font.color = "${font.color}";`);
  }
  
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (${params.slideIndex} < 0 || ${params.slideIndex} >= slides.items.length) {
    return {
      success: false,
      error: "幻灯片索引越界: ${params.slideIndex}"
    };
  }
  
  const slide = slides.items[${params.slideIndex}];
  const shape = slide.shapes.getItem("${params.tableId}");
  const table = shape.table;
  const cell = table.getCell(${params.row}, ${params.column});
  
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
export function generateDeleteTableCode(params: {
  slideIndex: number;
  tableId: string;
}): string {
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  if (${params.slideIndex} < 0 || ${params.slideIndex} >= slides.items.length) {
    return {
      success: false,
      error: "幻灯片索引越界: ${params.slideIndex}"
    };
  }
  
  const slide = slides.items[${params.slideIndex}];
  const shape = slide.shapes.getItem("${params.tableId}");
  shape.delete();
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}
