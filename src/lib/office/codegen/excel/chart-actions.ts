/**
 * Excel Chart Actions 代码生成器
 */

/**
 * 生成创建图表代码
 */
export function generateCreateCode(params: {
  dataRange: string;
  chartType: string;
  position?: { left: number; top: number };
  title?: string;
}): string {
  const position = params.position || { left: 200, top: 100 };
  
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("${params.dataRange}");
  
  // 创建图表
  const chart = sheet.charts.add(
    Excel.ChartType.${params.chartType},
    range,
    Excel.ChartSeriesBy.auto
  );
  
  chart.setPosition(null, null, ${position.left}, ${position.top});
  ${params.title ? `chart.title.text = "${params.title}";` : ''}
  
  chart.load("id");
  await context.sync();
  
  return {
    chartId: chart.id,
    success: true
  };
});`.trim();
}

/**
 * 生成更新图表代码
 */
export function generateUpdateCode(params: {
  chartId: string;
  dataRange?: string;
}): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const chart = sheet.charts.getItem("${params.chartId}");
  
  ${params.dataRange ? `
  const newRange = sheet.getRange("${params.dataRange}");
  chart.setData(newRange);
  ` : ''}
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成设置标题代码
 */
export function generateSetTitleCode(params: {
  chartId: string;
  title: string;
}): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const chart = sheet.charts.getItem("${params.chartId}");
  chart.title.text = "${params.title}";
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除图表代码
 */
export function generateDeleteCode(params: { chartId: string }): string {
  return `
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const chart = sheet.charts.getItem("${params.chartId}");
  chart.delete();
  
  await context.sync();
  
  return {
    success: true,
    deleted: "${params.chartId}"
  };
});`.trim();
}
