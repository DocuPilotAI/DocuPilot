/**
 * PowerPoint Shape Actions 代码生成器
 */

/**
 * 生成添加文本框代码
 */
export function generateAddTextCode(params: {
  slideIndex: number;
  text: string;
  position: any;
  format?: any;
}): string {
  const { left, top, width = 300, height = 100 } = params.position;
  const formatCode = params.format ? generateTextFormatCode(params.format, 'textBox') : '';
  
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
  const textBox = slide.shapes.addTextBox("${params.text.replace(/"/g, '\\"')}", {
    left: ${left},
    top: ${top},
    width: ${width},
    height: ${height}
  });
  
  textBox.load("id");
  ${formatCode}
  
  await context.sync();
  
  return {
    shapeId: textBox.id,
    success: true
  };
});`.trim();
}

/**
 * 生成添加图片代码
 */
export function generateAddImageCode(params: {
  slideIndex: number;
  base64Image: string;
  position: any;
}): string {
  const { left, top, width, height } = params.position;
  const sizeParams = width && height ? `, width: ${width}, height: ${height}` : '';
  
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
  
  // 移除 data:image/xxx;base64, 前缀（如果有）
  let imageData = "${params.base64Image}";
  if (imageData.includes(",")) {
    imageData = imageData.split(",")[1];
  }
  
  const image = slide.shapes.addImage({
    base64ImageString: imageData,
    left: ${left},
    top: ${top}${sizeParams}
  });
  
  image.load("id");
  await context.sync();
  
  return {
    shapeId: image.id,
    success: true
  };
});`.trim();
}

/**
 * 生成添加形状代码
 */
export function generateAddShapeCode(params: {
  slideIndex: number;
  shapeType: string;
  position: any;
  format?: any;
}): string {
  const { left, top, width = 200, height = 200 } = params.position;
  const shapeTypeMap: any = {
    rectangle: 'PowerPoint.ShapeType.rectangle',
    ellipse: 'PowerPoint.ShapeType.oval',
    triangle: 'PowerPoint.ShapeType.triangle',
    rightArrow: 'PowerPoint.ShapeType.rightArrow',
    leftArrow: 'PowerPoint.ShapeType.leftArrow',
    upArrow: 'PowerPoint.ShapeType.upArrow',
    downArrow: 'PowerPoint.ShapeType.downArrow',
    star: 'PowerPoint.ShapeType.star5'
  };
  
  const shapeType = shapeTypeMap[params.shapeType] || 'PowerPoint.ShapeType.rectangle';
  const formatCode = params.format ? generateShapeFormatCode(params.format, 'shape') : '';
  
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
  const shape = slide.shapes.addGeometricShape(${shapeType}, {
    left: ${left},
    top: ${top},
    width: ${width},
    height: ${height}
  });
  
  shape.load("id");
  ${formatCode}
  
  await context.sync();
  
  return {
    shapeId: shape.id,
    success: true
  };
});`.trim();
}

/**
 * 生成更新文本代码
 */
export function generateUpdateTextCode(params: {
  slideIndex: number;
  shapeId: string;
  text: string;
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
  const shape = slide.shapes.getItem("${params.shapeId}");
  const textFrame = shape.textFrame;
  textFrame.textRange.text = "${params.text.replace(/"/g, '\\"')}";
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成格式化形状代码
 */
export function generateFormatShapeCode(params: {
  slideIndex: number;
  shapeId: string;
  format: any;
}): string {
  const formatCode = generateShapeFormatCode(params.format, 'shape');
  
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
  const shape = slide.shapes.getItem("${params.shapeId}");
  
  ${formatCode}
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成移动形状代码
 */
export function generateMoveShapeCode(params: {
  slideIndex: number;
  shapeId: string;
  position: any;
}): string {
  const { left, top } = params.position;
  
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
  const shape = slide.shapes.getItem("${params.shapeId}");
  shape.left = ${left};
  shape.top = ${top};
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成删除形状代码
 */
export function generateDeleteShapeCode(params: {
  slideIndex: number;
  shapeId: string;
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
  const shape = slide.shapes.getItem("${params.shapeId}");
  shape.delete();
  
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}

/**
 * 生成文本格式化代码片段（辅助函数）
 */
function generateTextFormatCode(format: any, target: string): string {
  const formatters: string[] = [];
  
  if (format.fontSize) {
    formatters.push(`${target}.textFrame.textRange.font.size = ${format.fontSize};`);
  }
  if (format.bold !== undefined) {
    formatters.push(`${target}.textFrame.textRange.font.bold = ${format.bold};`);
  }
  if (format.italic !== undefined) {
    formatters.push(`${target}.textFrame.textRange.font.italic = ${format.italic};`);
  }
  if (format.underline !== undefined) {
    const underlineType = format.underline ? 'PowerPoint.FontUnderlineStyle.single' : 'PowerPoint.FontUnderlineStyle.none';
    formatters.push(`${target}.textFrame.textRange.font.underline = ${underlineType};`);
  }
  if (format.color) {
    formatters.push(`${target}.textFrame.textRange.font.color = "${format.color}";`);
  }
  if (format.fontName) {
    formatters.push(`${target}.textFrame.textRange.font.name = "${format.fontName}";`);
  }
  if (format.alignment) {
    formatters.push(`${target}.textFrame.textRange.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.${format.alignment.toLowerCase()};`);
  }
  if (format.verticalAlignment) {
    const valignMap: any = {
      Top: 'top',
      Middle: 'middle',
      Bottom: 'bottom'
    };
    formatters.push(`${target}.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.${valignMap[format.verticalAlignment] || 'top'};`);
  }
  
  return formatters.join('\n  ');
}

/**
 * 生成形状格式化代码片段（辅助函数）
 */
function generateShapeFormatCode(format: any, target: string): string {
  const formatters: string[] = [];
  
  if (format.fill?.color) {
    formatters.push(`${target}.fill.setSolidColor("${format.fill.color}");`);
  }
  
  if (format.line) {
    if (format.line.color) {
      formatters.push(`${target}.lineFormat.color = "${format.line.color}";`);
    }
    if (format.line.weight) {
      formatters.push(`${target}.lineFormat.weight = ${format.line.weight};`);
    }
  }
  
  return formatters.join('\n  ');
}
