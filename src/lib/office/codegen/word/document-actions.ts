/**
 * Word Document Actions 代码生成器
 */

/**
 * 生成读取文档代码
 */
export function generateReadCode(params: {
  includeFormat?: boolean;
  maxLength?: number;
}): string {
  const maxLen = params.maxLength || 10000;
  
  return `
Word.run(async (context) => {
  const body = context.document.body;
  body.load("text");
  await context.sync();
  
  let text = body.text;
  if (text.length > ${maxLen}) {
    text = text.substring(0, ${maxLen}) + "... (truncated)";
  }
  
  // 获取段落数和单词数
  const paragraphs = body.paragraphs;
  paragraphs.load("items");
  await context.sync();
  
  return {
    text: text,
    paragraphCount: paragraphs.items.length,
    wordCount: text.split(/\\s+/).filter(word => word.length > 0).length
  };
});`.trim();
}

/**
 * 生成读取选中文本代码
 */
export function generateReadSelectionCode(): string {
  return `
Word.run(async (context) => {
  const selection = context.document.getSelection();
  selection.load("text, style");
  selection.font.load("name, size, bold, italic, color");
  await context.sync();
  
  return {
    text: selection.text,
    style: selection.style,
    font: {
      name: selection.font.name,
      size: selection.font.size,
      bold: selection.font.bold,
      italic: selection.font.italic,
      color: selection.font.color
    }
  };
});`.trim();
}

/**
 * 生成搜索文本代码
 */
export function generateSearchCode(params: {
  searchText: string;
  matchCase?: boolean;
  matchWholeWord?: boolean;
}): string {
  return `
Word.run(async (context) => {
  const results = context.document.body.search("${params.searchText.replace(/"/g, '\\"')}", {
    matchCase: ${params.matchCase || false},
    matchWholeWord: ${params.matchWholeWord || false}
  });
  results.load("items");
  await context.sync();
  
  const matches = results.items.map((item, index) => {
    item.load("text");
    return {
      text: item.text,
      context: "",  // 简化版本，不获取上下文
      index: index
    };
  });
  
  await context.sync();
  
  return {
    matches: matches.map(m => ({
      text: m.text,
      context: m.context
    })),
    count: results.items.length
  };
});`.trim();
}

/**
 * 生成替换文本代码
 */
export function generateReplaceCode(params: {
  searchText: string;
  replaceText: string;
  matchCase?: boolean;
  replaceAll?: boolean;
}): string {
  return `
Word.run(async (context) => {
  const results = context.document.body.search("${params.searchText.replace(/"/g, '\\"')}", {
    matchCase: ${params.matchCase || false}
  });
  results.load("items");
  await context.sync();
  
  let replacedCount = 0;
  
  ${params.replaceAll ? `
  // 替换所有匹配项
  for (const item of results.items) {
    item.insertText("${params.replaceText.replace(/"/g, '\\"')}", Word.InsertLocation.replace);
    replacedCount++;
  }
  ` : `
  // 只替换第一个匹配项
  if (results.items.length > 0) {
    results.items[0].insertText("${params.replaceText.replace(/"/g, '\\"')}", Word.InsertLocation.replace);
    replacedCount = 1;
  }
  `}
  
  await context.sync();
  
  return {
    replacedCount: replacedCount
  };
});`.trim();
}

/**
 * 生成清除文档代码
 */
export function generateClearCode(): string {
  return `
Word.run(async (context) => {
  const body = context.document.body;
  body.clear();
  await context.sync();
  
  return {
    success: true
  };
});`.trim();
}
