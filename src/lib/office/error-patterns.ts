/**
 * 错误模式映射
 * 提供常见 Office.js 错误的修复建议和示例
 */

import { ExecutionError } from './code-executor';

export interface ErrorPattern {
  hint: string;
  examples: string[];
  commonCauses?: string[];
}

/**
 * Office.js 常见错误模式及修复建议
 */
export const ERROR_FIX_HINTS: Record<string, ErrorPattern> = {
  InvalidArgument: {
    hint: "参数验证失败，检查参数名称、类型和取值范围",
    examples: [
      "使用正确的 InsertLocation 枚举值（如 'Before'、'After'、'Start'、'End'）",
      "确保单元格地址格式正确（如 'A1'、'B2:D4'）",
      "检查数组索引是否越界",
      "验证参数类型是否匹配（字符串、数字、布尔值）",
      "确保必填参数不为 null 或 undefined"
    ],
    commonCauses: [
      "参数拼写错误",
      "使用了错误的枚举值",
      "参数类型不匹配",
      "缺少必需的参数"
    ]
  },

  InvalidReference: {
    hint: "引用的对象不存在或无效，需要先检查对象是否存在",
    examples: [
      "使用 getItemOrNullObject() 代替 getItem()，然后检查 isNullObject",
      "在操作前调用 load() 和 await context.sync()",
      "验证工作表、单元格、命名项是否存在",
      "确保在创建对象后再进行引用",
      "检查名称是否正确（区分大小写）"
    ],
    commonCauses: [
      "引用不存在的工作表或命名范围",
      "对象已被删除",
      "名称拼写错误",
      "在同步前访问对象属性"
    ]
  },

  ApiNotFound: {
    hint: "API 在当前 Office 版本或平台不可用",
    examples: [
      "使用 Office.context.requirements.isSetSupported() 检查 API 可用性",
      "提供降级方案或替代实现",
      "使用更通用、兼容性更好的 API",
      "检查 Office 版本要求（如需要 Office 2016 或更高版本）",
      "针对不同平台提供不同实现（Web、Desktop、Mac）"
    ],
    commonCauses: [
      "使用了新版本才支持的 API",
      "在不支持的平台上运行",
      "API 拼写错误",
      "缺少必要的权限集"
    ]
  },

  GeneralException: {
    hint: "Office.js 内部错误或操作冲突",
    examples: [
      "简化操作步骤，避免一次性执行过多操作",
      "分批执行，每批后调用 context.sync()",
      "添加适当的延迟（setTimeout）",
      "检查是否有并发操作冲突",
      "确保操作顺序正确（先创建后修改）"
    ],
    commonCauses: [
      "一次加载过多数据",
      "操作顺序不当",
      "并发操作冲突",
      "Office 应用内部状态异常"
    ]
  },

  NetworkError: {
    hint: "网络连接问题或超时",
    examples: [
      "检查网络连接是否正常",
      "增加超时时间",
      "添加重试逻辑",
      "减少单次请求的数据量",
      "使用本地缓存"
    ],
    commonCauses: [
      "网络不稳定",
      "请求超时",
      "服务器无响应",
      "数据量过大"
    ]
  },

  UnknownError: {
    hint: "未知错误，需要进一步诊断",
    examples: [
      "检查控制台中的详细错误信息",
      "验证代码语法是否正确",
      "确保所有异步操作都使用了 await",
      "添加 try-catch 捕获具体错误",
      "简化代码逐步定位问题"
    ],
    commonCauses: [
      "JavaScript 语法错误",
      "未处理的异步操作",
      "变量未定义",
      "类型转换错误"
    ]
  }
};

/**
 * 获取错误类型的修复建议
 * 
 * @param errorType - 错误类型
 * @returns 格式化的修复建议文本
 */
export function getFixHint(errorType: string): string {
  const pattern = ERROR_FIX_HINTS[errorType];
  
  if (!pattern) {
    return `
## 修复建议

未找到针对此错误类型的具体建议，请尝试：
- 检查错误消息中的具体提示
- 查阅 Office.js 文档
- 简化代码逐步定位问题
- 添加更详细的错误处理
`;
  }

  let hint = `
## 修复建议

**问题描述**: ${pattern.hint}

**常见原因**:
${pattern.commonCauses ? pattern.commonCauses.map(cause => `- ${cause}`).join('\n') : '- 参考错误消息中的具体信息'}

**解决方案**:
${pattern.examples.map(example => `- ${example}`).join('\n')}
`;

  return hint;
}

/**
 * 根据错误类型和消息推断可能的原因
 * 
 * @param error - 执行错误对象
 * @returns 可能的原因列表
 */
export function inferErrorCauses(error: ExecutionError): string[] {
  const causes: string[] = [];
  const message = error.message.toLowerCase();

  // InvalidArgument 相关
  if (error.type === 'InvalidArgument') {
    if (message.includes('name') || message.includes('名称')) {
      causes.push('参数名称可能拼写错误或不存在');
    }
    if (message.includes('range') || message.includes('地址')) {
      causes.push('单元格地址格式可能不正确');
    }
    if (message.includes('value') || message.includes('值')) {
      causes.push('参数值可能超出允许范围');
    }
  }

  // InvalidReference 相关
  if (error.type === 'InvalidReference') {
    if (message.includes('worksheet') || message.includes('工作表')) {
      causes.push('引用的工作表可能不存在');
    }
    if (message.includes('null') || message.includes('undefined')) {
      causes.push('对象可能未正确初始化');
    }
  }

  // ApiNotFound 相关
  if (error.type === 'ApiNotFound') {
    causes.push('当前 Office 版本可能不支持此 API');
    causes.push('需要检查 API 的平台兼容性');
  }

  // 如果没有特定原因，返回通用建议
  if (causes.length === 0) {
    const pattern = ERROR_FIX_HINTS[error.type];
    if (pattern?.commonCauses) {
      causes.push(...pattern.commonCauses);
    }
  }

  return causes;
}

/**
 * 生成代码修复建议（基于错误类型）
 * 
 * @param error - 执行错误对象
 * @param code - 失败的代码
 * @returns 具体的代码修复建议
 */
export function generateCodeFixSuggestions(error: ExecutionError, code: string): string[] {
  const suggestions: string[] = [];

  // InvalidArgument - 检查常见问题
  if (error.type === 'InvalidArgument') {
    if (code.includes('.getItem(')) {
      suggestions.push('考虑使用 getItemOrNullObject() 替代 getItem()');
    }
    if (code.includes('InsertLocation')) {
      suggestions.push('检查 InsertLocation 枚举值是否正确（使用 "Before"、"After" 等字符串）');
    }
  }

  // InvalidReference - 检查对象引用
  if (error.type === 'InvalidReference') {
    if (!code.includes('.load(')) {
      suggestions.push('在访问属性前添加 load() 和 sync()');
    }
    if (!code.includes('getItemOrNullObject')) {
      suggestions.push('使用 getItemOrNullObject() 检查对象是否存在');
    }
    if (!code.includes('isNullObject')) {
      suggestions.push('添加 isNullObject 检查，确保对象有效');
    }
  }

  // ApiNotFound - 检查 API 兼容性
  if (error.type === 'ApiNotFound') {
    suggestions.push('添加 isSetSupported() 检查，提供降级方案');
    suggestions.push('查阅文档确认 API 的版本要求');
  }

  // GeneralException - 简化操作
  if (error.type === 'GeneralException') {
    const syncCount = (code.match(/context\.sync\(\)/g) || []).length;
    if (syncCount === 0) {
      suggestions.push('添加 context.sync() 确保操作同步');
    }
    if (code.length > 500) {
      suggestions.push('将复杂操作拆分为多个小步骤');
    }
  }

  // 通用建议
  if (!code.includes('try') || !code.includes('catch')) {
    suggestions.push('添加 try-catch 错误处理');
  }

  return suggestions;
}
