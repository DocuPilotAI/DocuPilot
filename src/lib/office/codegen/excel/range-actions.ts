/**
 * Excel Range Actions 代码生成器
 */

/**
 * 生成读取区域代码
 */
export function generateReadCode(params: {
  address: string;
  includeFormulas?: boolean;
  includeFormat?: boolean;
}): string {
  const loadProps = ["values", "address", "rowCount", "columnCount"];
  if (params.includeFormulas) loadProps.push("formulas");
  if (params.includeFormat) loadProps.push("numberFormat");
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.address}");
  range.load(${JSON.stringify(loadProps)});
  await context.sync();
  
  return {
    address: range.address,
    values: range.values,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    ${params.includeFormulas ? 'formulas: range.formulas,' : ''}
    ${params.includeFormat ? 'numberFormat: range.numberFormat' : ''}
  };
});`.trim();
}

/**
 * 生成读取选中区域代码
 */
export function generateReadSelectionCode(): string {
  return `
Excel.run(async (context) => {
  const range = context.workbook.getSelectedRange();
  range.load(["values", "address", "rowCount", "columnCount"]);
  await context.sync();
  
  return {
    address: range.address,
    values: range.values,
    rowCount: range.rowCount,
    columnCount: range.columnCount
  };
});`.trim();
}

/**
 * 生成写入数据代码
 */
export function generateWriteCode(params: {
  address: string;
  values: any[][];
  autoExpand?: boolean;
}): string {
  const valuesJson = JSON.stringify(params.values);
  const autoExpand = params.autoExpand !== false;
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  let range = sheet.getRange("${params.address}");
  
  ${autoExpand ? `
  // 自动扩展区域以匹配数据大小
  const dataRows = ${params.values.length};
  const dataCols = ${params.values[0]?.length || 0};
  if (dataRows > 0 && dataCols > 0) {
    range = range.getResizedRange(dataRows - 1, dataCols - 1);
  }
  ` : ''}
  
  range.values = ${valuesJson};
  await context.sync();
  
  range.load("address");
  await context.sync();
  
  return {
    success: true,
    writtenRange: range.address
  };
});`.trim();
}

/**
 * 生成格式化代码
 */
export function generateFormatCode(params: {
  address: string;
  format: any;
}): string {
  const formatters: string[] = [];
  
  if (params.format.font) {
    const font = params.format.font;
    if (font.name) formatters.push(`range.format.font.name = "${font.name}";`);
    if (font.size) formatters.push(`range.format.font.size = ${font.size};`);
    if (font.bold !== undefined) formatters.push(`range.format.font.bold = ${font.bold};`);
    if (font.italic !== undefined) formatters.push(`range.format.font.italic = ${font.italic};`);
    if (font.underline !== undefined) formatters.push(`range.format.font.underline = ${font.underline ? '"Single"' : '"None"'};`);
    if (font.color) formatters.push(`range.format.font.color = "${font.color}";`);
  }
  
  if (params.format.fill?.color) {
    formatters.push(`range.format.fill.color = "${params.format.fill.color}";`);
  }
  
  if (params.format.borders) {
    const borders = params.format.borders;
    if (borders.style || borders.color || borders.weight) {
      // 设置所有边框
      formatters.push(`const borderEdges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];`);
      formatters.push(`borderEdges.forEach(edge => {`);
      if (borders.style) formatters.push(`  range.format.borders.getItem(edge).style = Excel.BorderLineStyle.${borders.style};`);
      if (borders.color) formatters.push(`  range.format.borders.getItem(edge).color = "${borders.color}";`);
      if (borders.weight) formatters.push(`  range.format.borders.getItem(edge).weight = Excel.BorderWeight.${borders.weight};`);
      formatters.push(`});`);
    }
  }
  
  if (params.format.numberFormat) {
    formatters.push(`range.numberFormat = "${params.format.numberFormat}";`);
  }
  
  if (params.format.horizontalAlignment) {
    formatters.push(`range.format.horizontalAlignment = Excel.HorizontalAlignment.${params.format.horizontalAlignment.toLowerCase()};`);
  }
  
  if (params.format.verticalAlignment) {
    formatters.push(`range.format.verticalAlignment = Excel.VerticalAlignment.${params.format.verticalAlignment.toLowerCase()};`);
  }
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.address}");
  
  ${formatters.join('\n  ')}
  
  await context.sync();
  
  return { success: true };
});`.trim();
}

/**
 * 生成清除代码
 */
export function generateClearCode(params: { address: string; applyTo?: 'contents' | 'formats' | 'all' }): string {
  const clearType = params.applyTo || 'all';
  const clearMethod = clearType === 'contents' ? 'clear(Excel.ClearApplyTo.contents)' :
                      clearType === 'formats' ? 'clear(Excel.ClearApplyTo.formats)' :
                      'clear(Excel.ClearApplyTo.all)';
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.address}");
  range.${clearMethod};
  await context.sync();
  
  return { success: true };
});`.trim();
}

/**
 * 生成复制代码
 */
export function generateCopyCode(params: { source: string; destination: string; copyType?: 'all' | 'values' | 'formats' }): string {
  const copyType = params.copyType || 'all';
  const copyMethod = copyType === 'values' ? 'Excel.RangeCopyType.values' :
                      copyType === 'formats' ? 'Excel.RangeCopyType.formats' :
                      'Excel.RangeCopyType.all';
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const sourceRange = sheet.getRange("${params.source}");
  const destinationRange = sheet.getRange("${params.destination}");
  
  sourceRange.copyFrom(destinationRange, ${copyMethod});
  await context.sync();
  
  return { success: true };
});`.trim();
}

/**
 * 生成插入/删除单元格代码
 */
export function generateInsertDeleteCode(
  action: 'insert' | 'delete',
  params: { address: string; shift: 'down' | 'right' | 'up' | 'left' }
): string {
  const shiftMapping = {
    down: 'Excel.InsertShiftDirection.down',
    right: 'Excel.InsertShiftDirection.right',
    up: 'Excel.DeleteShiftDirection.up',
    left: 'Excel.DeleteShiftDirection.left'
  };
  
  const shiftValue = shiftMapping[params.shift];
  const operation = action === 'insert' ? `insert(${shiftValue})` : `delete(${shiftValue})`;
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.address}");
  range.${operation};
  await context.sync();
  
  return { success: true };
});`.trim();
}
