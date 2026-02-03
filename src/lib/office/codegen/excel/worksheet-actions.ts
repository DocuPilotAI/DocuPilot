/**
 * Excel Worksheet Actions 代码生成器
 */

/**
 * 生成列出工作表代码
 */
export function generateListCode(): string {
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name, items/id, items/position, items/visibility");
  await context.sync();
  
  const sheetInfos = sheets.items.map(sheet => ({
    name: sheet.name,
    id: sheet.id,
    position: sheet.position,
    visible: sheet.visibility === Excel.SheetVisibility.visible
  }));
  
  return {
    sheets: sheetInfos,
    count: sheetInfos.length
  };
});`.trim();
}

/**
 * 生成添加工作表代码
 */
export function generateAddCode(params: {
  name: string;
  position?: 'start' | 'end' | 'before' | 'after';
  referenceSheet?: string;
}): string {
  if (params.position === 'before' || params.position === 'after') {
    // 需要参考工作表
    const positionType = params.position === 'before' ? 'before' : 'after';
    return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  
  // 检查工作表是否已存在
  const existingSheet = sheets.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (!existingSheet.isNullObject) {
    return {
      name: "${params.name}",
      id: existingSheet.id,
      success: false,
      error: "工作表已存在"
    };
  }
  
  // 获取参考工作表
  const refSheet = sheets.getItemOrNullObject("${params.referenceSheet || ''}");
  await context.sync();
  
  if (refSheet.isNullObject) {
    return {
      success: false,
      error: "参考工作表不存在: ${params.referenceSheet || ''}"
    };
  }
  
  // 添加工作表
  const newSheet = sheets.add("${params.name}", Excel.WorksheetPositionType.${positionType}, refSheet);
  newSheet.load("name, id");
  await context.sync();
  
  return {
    name: newSheet.name,
    id: newSheet.id,
    success: true
  };
});`.trim();
  } else {
    // 简单添加到开始或结束
    const positionType = params.position === 'start' ? 'start' : 'end';
    return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  
  // 检查工作表是否已存在
  const existingSheet = sheets.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (!existingSheet.isNullObject) {
    return {
      name: "${params.name}",
      id: existingSheet.id,
      success: false,
      error: "工作表已存在"
    };
  }
  
  // 添加工作表
  const newSheet = sheets.add("${params.name}", Excel.WorksheetPositionType.${positionType});
  newSheet.load("name, id");
  await context.sync();
  
  return {
    name: newSheet.name,
    id: newSheet.id,
    success: true
  };
});`.trim();
  }
}

/**
 * 生成删除工作表代码
 */
export function generateDeleteCode(params: { name: string }): string {
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (sheet.isNullObject) {
    return {
      success: false,
      error: "工作表不存在: ${params.name}"
    };
  }
  
  sheet.delete();
  await context.sync();
  
  return {
    success: true,
    deleted: "${params.name}"
  };
});`.trim();
}

/**
 * 生成重命名工作表代码
 */
export function generateRenameCode(params: { oldName: string; newName: string }): string {
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.getItemOrNullObject("${params.oldName}");
  await context.sync();
  
  if (sheet.isNullObject) {
    return {
      success: false,
      error: "工作表不存在: ${params.oldName}"
    };
  }
  
  sheet.name = "${params.newName}";
  await context.sync();
  
  return {
    success: true,
    oldName: "${params.oldName}",
    newName: "${params.newName}"
  };
});`.trim();
}

/**
 * 生成检查工作表存在代码
 */
export function generateExistsCode(params: { name: string }): string {
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.getItemOrNullObject("${params.name}");
  await context.sync();
  
  return {
    exists: !sheet.isNullObject,
    name: "${params.name}"
  };
});`.trim();
}

/**
 * 生成激活工作表代码
 */
export function generateActivateCode(params: { name: string }): string {
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sheet = sheets.getItemOrNullObject("${params.name}");
  await context.sync();
  
  if (sheet.isNullObject) {
    return {
      success: false,
      error: "工作表不存在: ${params.name}"
    };
  }
  
  sheet.activate();
  await context.sync();
  
  return {
    success: true,
    activated: "${params.name}"
  };
});`.trim();
}

/**
 * 生成复制工作表代码
 */
export function generateCopyCode(params: { sourceName: string; newName: string; position?: 'before' | 'after' }): string {
  const positionType = params.position === 'before' ? 'before' : 'after';
  
  return `
Excel.run(async (context) => {
  const sheets = context.workbook.worksheets;
  const sourceSheet = sheets.getItemOrNullObject("${params.sourceName}");
  await context.sync();
  
  if (sourceSheet.isNullObject) {
    return {
      success: false,
      error: "源工作表不存在: ${params.sourceName}"
    };
  }
  
  // 复制工作表
  const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.${positionType}, sourceSheet);
  copiedSheet.name = "${params.newName}";
  copiedSheet.load("name, id");
  await context.sync();
  
  return {
    success: true,
    newSheetName: copiedSheet.name,
    newSheetId: copiedSheet.id
  };
});`.trim();
}
