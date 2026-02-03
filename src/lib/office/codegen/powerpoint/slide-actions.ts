/**
 * PowerPoint Slide Actions 代码生成器
 */

/**
 * 生成列出幻灯片代码
 */
export function generateListCode(): string {
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items");
  await context.sync();
  
  const slideInfos = slides.items.map((slide, index) => ({
    index: index,
    id: slide.id
  }));
  
  return {
    slideCount: slides.items.length,
    slides: slideInfos
  };
});`.trim();
}

/**
 * 生成读取幻灯片内容代码
 */
export function generateReadCode(params: { slideIndex: number }): string {
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
  const shapes = slide.shapes;
  shapes.load("items");
  await context.sync();
  
  const texts = [];
  const shapeInfos = [];
  
  for (const shape of shapes.items) {
    shape.load("id, type");
    if (shape.type === PowerPoint.ShapeType.textBox || shape.type === PowerPoint.ShapeType.geometricShape) {
      const textFrame = shape.textFrame;
      textFrame.load("textRange");
      await context.sync();
      
      texts.push({
        content: textFrame.textRange.text,
        shapeType: shape.type
      });
    }
    
    shapeInfos.push({
      id: shape.id,
      type: shape.type
    });
  }
  
  await context.sync();
  
  return {
    texts: texts,
    shapes: shapeInfos,
    shapeCount: shapes.items.length
  };
});`.trim();
}

/**
 * 生成添加幻灯片代码
 */
export function generateAddCode(params: { layout?: string; insertAfter?: number }): string {
  const layoutType = params.layout || 'Blank';
  
  return `
PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  
  // 添加新幻灯片
  slides.add();
  await context.sync();
  
  // 重新加载以获取新幻灯片
  slides.load("items");
  await context.sync();
  
  const newSlideIndex = slides.items.length - 1;
  
  return {
    slideIndex: newSlideIndex,
    success: true
  };
});`.trim();
}

/**
 * 生成删除幻灯片代码
 */
export function generateDeleteCode(params: { slideIndex: number }): string {
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
  
  slides.items[${params.slideIndex}].delete();
  await context.sync();
  
  return {
    success: true,
    deleted: ${params.slideIndex}
  };
});`.trim();
}

/**
 * 生成复制幻灯片代码
 */
export function generateDuplicateCode(params: { slideIndex: number }): string {
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
  
  // PowerPoint.js 不直接支持复制幻灯片
  // 作为替代，我们可以添加新幻灯片并提示用户
  slides.add();
  await context.sync();
  
  slides.load("items");
  await context.sync();
  
  const newSlideIndex = slides.items.length - 1;
  
  return {
    newSlideIndex: newSlideIndex,
    success: true,
    note: "已添加新幻灯片，PowerPoint.js暂不支持自动复制内容"
  };
});`.trim();
}

/**
 * 生成移动幻灯片代码
 */
export function generateMoveCode(params: { fromIndex: number; toIndex: number }): string {
  return `
PowerPoint.run(async (context) => {
  // PowerPoint.js 暂不支持直接移动幻灯片
  // 返回提示信息
  return {
    success: false,
    error: "PowerPoint.js API 暂不支持移动幻灯片功能"
  };
});`.trim();
}
