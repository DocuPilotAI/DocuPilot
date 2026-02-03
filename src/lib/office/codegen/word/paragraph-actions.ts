/**
 * Word Paragraph Actions 代码生成器
 */

/**
 * 生成插入段落代码
 */
export function generateInsertCode(params: {
  text: string;
  location: string;
  format?: any;
}): string {
  const formatCode = params.format ? generateFormatCode(params.format, 'paragraph') : '';
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const paragraph = body.insertParagraph("${params.text.replace(/"/g, '\\"')}", Word.InsertLocation.${params.location.toLowerCase()});
  
  ${formatCode}
  
  await context.sync();
  
  // 获取段落索引
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  let paragraphIndex = -1;
  for (let i = 0; i < paragraphs.items.length; i++) {
    if (paragraphs.items[i] === paragraph) {
      paragraphIndex = i;
      break;
    }
  }
  
  return {
    success: true,
    paragraphIndex: paragraphIndex
  };
});`.trim();
}

/**
 * 生成在指定位置插入段落代码
 */
export function generateInsertAtCode(params: {
  text: string;
  index: number;
  location: string;
  format?: any;
}): string {
  const formatCode = params.format ? generateFormatCode(params.format, 'paragraph') : '';
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  if (${params.index} < 0 || ${params.index} >= paragraphs.items.length) {
    return {
      success: false,
      error: "段落索引越界: ${params.index}"
    };
  }
  
  const targetParagraph = paragraphs.items[${params.index}];
  const paragraph = targetParagraph.insertParagraph("${params.text.replace(/"/g, '\\"')}", Word.InsertLocation.${params.location.toLowerCase()});
  
  ${formatCode}
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成格式化段落代码
 */
export function generateFormatParagraphCode(params: {
  index: number;
  format: any;
}): string {
  const formatCode = generateFormatCode(params.format, 'paragraph');
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  if (${params.index} < 0 || ${params.index} >= paragraphs.items.length) {
    return {
      success: false,
      error: "段落索引越界: ${params.index}"
    };
  }
  
  const paragraph = paragraphs.items[${params.index}];
  
  ${formatCode}
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除段落代码
 */
export function generateDeleteCode(params: { index: number }): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  if (${params.index} < 0 || ${params.index} >= paragraphs.items.length) {
    return {
      success: false,
      error: "段落索引越界: ${params.index}"
    };
  }
  
  paragraphs.items[${params.index}].delete();
  await context.sync();
  
  return {
    success: true,
    deleted: ${params.index}
  };
});`.trim();
}

/**
 * 生成获取段落信息代码
 */
export function generateGetCode(params: { index: number }): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  if (${params.index} < 0 || ${params.index} >= paragraphs.items.length) {
    return {
      success: false,
      error: "段落索引越界: ${params.index}"
    };
  }
  
  const paragraph = paragraphs.items[${params.index}];
  paragraph.load("text, style");
  paragraph.font.load("name, size, bold, italic, color");
  await context.sync();
  
  return {
    text: paragraph.text,
    style: paragraph.style,
    font: {
      name: paragraph.font.name,
      size: paragraph.font.size,
      bold: paragraph.font.bold,
      italic: paragraph.font.italic,
      color: paragraph.font.color
    }
  };
});`.trim();
}

/**
 * 生成格式化代码片段（辅助函数）
 */
function generateFormatCode(format: any, target: 'paragraph' | 'range' = 'paragraph'): string {
  const formatters: string[] = [];
  
  if (format.style) {
    formatters.push(`${target}.style = "${format.style}";`);
  }
  
  if (format.alignment) {
    formatters.push(`${target}.alignment = Word.Alignment.${format.alignment.toLowerCase()};`);
  }
  
  if (format.font) {
    const font = format.font;
    if (font.name) formatters.push(`${target}.font.name = "${font.name}";`);
    if (font.size) formatters.push(`${target}.font.size = ${font.size};`);
    if (font.bold !== undefined) formatters.push(`${target}.font.bold = ${font.bold};`);
    if (font.italic !== undefined) formatters.push(`${target}.font.italic = ${font.italic};`);
    if (font.underline !== undefined) {
      formatters.push(`${target}.font.underline = Word.UnderlineType.${font.underline ? 'single' : 'none'};`);
    }
    if (font.color) formatters.push(`${target}.font.color = "${font.color}";`);
  }
  
  if (format.spacing) {
    const spacing = format.spacing;
    if (spacing.before !== undefined) formatters.push(`${target}.spaceAfter = ${spacing.before};`);
    if (spacing.after !== undefined) formatters.push(`${target}.spaceBefore = ${spacing.after};`);
    if (spacing.line !== undefined) formatters.push(`${target}.lineSpacing = ${spacing.line};`);
  }
  
  return formatters.join('\n  ');
}
