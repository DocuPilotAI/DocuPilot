#!/usr/bin/env node
/**
 * 从 TOOLS.md 文件生成测试用例
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

// 获取当前文件的目录路径 (ESM 模块中的 __dirname 替代)
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface TestCase {
  id: string;
  category: string;
  name: string;
  description: string;
  userInput: string;
  expectedCode?: string;
  expectedBehavior: string;
  validationSteps: string[];
  toolsTemplate: string;
  priority?: 'high' | 'medium' | 'low';
  tags?: string[];
}

interface TestSuite {
  id: string;
  name: string;
  hostType: 'excel' | 'word' | 'powerpoint';
  testCases: TestCase[];
  metadata: {
    version: string;
    generatedAt: string;
    sourceFile: string;
  };
}

/**
 * 解析 TOOLS.md 文件,提取代码模板
 */
function parseToolsFile(content: string): Array<{ category: string; name: string; code: string }> {
  const templates: Array<{ category: string; name: string; code: string }> = [];
  
  // 匹配标题和代码块
  const sections = content.split(/^##\s+/m).filter(s => s.trim());
  
  for (const section of sections) {
    const lines = section.split('\n');
    const category = lines[0].trim().replace(/模板$/, '').trim();
    
    // 查找所有 ### 标题和对应的代码块
    let currentName = '';
    let collectingCode = false;
    let codeLines: string[] = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // ### 标题
      if (line.match(/^###\s+/)) {
        // 保存前一个模板
        if (currentName && codeLines.length > 0) {
          templates.push({
            category,
            name: currentName,
            code: codeLines.join('\n').trim(),
          });
          codeLines = [];
        }
        currentName = line.replace(/^###\s+/, '').trim();
        collectingCode = false;
      }
      // 代码块开始
      else if (line.trim() === '```javascript' || line.trim() === '```js') {
        collectingCode = true;
      }
      // 代码块结束
      else if (line.trim() === '```' && collectingCode) {
        collectingCode = false;
      }
      // 收集代码
      else if (collectingCode) {
        codeLines.push(line);
      }
    }
    
    // 保存最后一个模板
    if (currentName && codeLines.length > 0) {
      templates.push({
        category,
        name: currentName,
        code: codeLines.join('\n').trim(),
      });
    }
  }
  
  return templates;
}

/**
 * 根据模板生成测试用例
 */
function generateTestCase(
  template: { category: string; name: string; code: string },
  index: number,
  hostType: 'excel' | 'word' | 'powerpoint'
): TestCase {
  const id = `${hostType}-${String(index + 1).padStart(3, '0')}-${template.name.toLowerCase().replace(/\s+/g, '-')}`;
  
  // 根据模板名称生成用户输入
  const userInput = generateUserInput(template.name, template.category, hostType);
  
  // 生成期望行为
  const expectedBehavior = generateExpectedBehavior(template.name, template.category);
  
  // 生成验证步骤
  const validationSteps = generateValidationSteps(template.name, template.category);
  
  // 确定优先级
  const priority = determinePriority(template.category, template.name);
  
  return {
    id,
    category: template.category,
    name: template.name,
    description: userInput,
    userInput,
    expectedCode: template.code,
    expectedBehavior,
    validationSteps,
    toolsTemplate: `${template.category} > ${template.name}`,
    priority,
  };
}

/**
 * 生成用户输入文本
 */
function generateUserInput(name: string, category: string, hostType: string): string {
  const inputs: Record<string, string> = {
    // Excel 工作表管理
    '创建工作表': '创建一个名为 NewSheet 的工作表',
    '创建工作表（带验证）': '创建一个名为 NewSheet 的工作表，如果已存在则使用现有的',
    '重命名工作表': '将当前工作表重命名为 RenamedSheet',
    '删除工作表': '删除名为 SheetToDelete 的工作表',
    '删除工作表（带验证）': '检查并删除名为 SheetToDelete 的工作表，如果不存在则给出提示',
    '复制工作表': '复制当前工作表',
    
    // Excel 条件格式
    '单元格值条件格式': '高亮显示所有大于 100 的单元格',
    '数据条格式': '为 B 列添加数据条显示',
    '图标集格式': '用箭头图标集显示 C 列数据趋势',
    '预设条件格式（高于平均值）': '高亮显示高于平均值的单元格',
    '前/后N项格式': '高亮显示前 10 名的数据',
    '自定义公式条件格式': '用公式设置条件：如果当前单元格大于左边的单元格则显示为绿色',
    
    // Excel 事件处理
    '监听单元格数据变化': '监听工作表的数据变化，当单元格内容改变时输出日志',
    '监听选择区域变化': '监听用户选择的单元格变化',
    '工作表激活事件': '监听工作表被激活的事件',
    '计算完成事件': '监听工作表计算完成事件',
    
    // Excel 形状操作
    '添加矩形形状': '在工作表中添加一个蓝色矩形',
    '插入图片': '在工作表中插入一张图片',
    '添加线条': '添加一条红色箭头线',
    '创建文本框': '创建一个包含"重要提示"的文本框',
    '形状分组': '将 Shape1 和 Shape2 组合为一个组',
    
    // Excel 备注
    '添加单元格备注': '在 A1 单元格添加备注"这是重要数据"',
    '修改备注内容': '修改 A1 单元格的备注内容',
    '设置备注可见性': '设置 A1 的备注始终显示',
    '删除备注': '删除 A2 单元格的备注',
    
    // Excel 范围高级操作
    '复制区域到新位置': '将 A1:C5 复制到 E1',
    '复制时跳过空白单元格': '复制 A1:C3 到 D1，跳过空白单元格',
    '移动区域': '将 A1:C5 移动到 G1',
    '插入单元格': '在 B2 插入单元格，其他单元格向下移动',
    '删除单元格': '删除 C3:C5，其他单元格向上移动',
    '清除区域内容': '清除 A1:D10 区域的所有内容',
    '删除重复值': '删除 A1:B20 区域中基于第1和第2列的重复行',
    '行列分组': '对第 2 到第 5 行进行分组',
    '取消分组': '取消第 2 到第 5 行的分组',
    
    // Excel 复选框
    '添加复选框': '将 B2:B10 的布尔值转换为复选框',
    '读取复选框状态': '读取 B2:B10 复选框的选中状态',
    '移除复选框': '移除 B2:B10 的复选框，还原为布尔值',
    
    // Excel 数据操作
    '读取选中区域': '读取当前选中的单元格数据',
    '读取指定区域': '读取 A1:D10 区域的数据',
    '读取多个不连续区域': '读取 B3, F3, J3 这三个单元格的值',
    '读取多个不连续区域（重要！）': '分别读取 B3, F3, J3 这三个单元格的值',
    '写入单个值': '在 A1 单元格写入数字 100',
    '写入数组数据': '在 A1:C3 区域写入示例数据',
    '写入多个不连续区域': '在 B3, F3, J3 单元格分别写入 10, 20, 30',
    
    // Excel 表格
    'excel:创建表格': '将 A1:D10 区域创建为表格',
    '创建表格': '创建一个名为 SalesTable 的表格，包含 Date, Product, Category, Amount 列',
    '排序表格': '按第四列 Amount 降序排序 SalesTable 表格',
    '筛选表格': '筛选 SalesTable 表格第三列 Category 为 Electronics 的数据',
    
    // Excel 数据透视表
    '创建透视表': '基于 A1:D5 数据在 F1 位置创建数据透视表，按 Category 和 Product 汇总 Amount',
    '添加列层次结构和筛选器': '为 PivotTable1 添加 Region 列字段和 Category 筛选字段',
    '数据透视表日期筛选': '筛选 PivotTable1 的日期字段，只显示 2020-08-01 之后的数据',
    '数据透视表标签筛选': '筛选 PivotTable1 的 Category 字段，排除以"电子"开头的项',
    '数据透视表数值筛选': '筛选 PivotTable1 的 Product 字段，只显示 Amount 大于 500 的项',
    '清除透视表筛选器': '清除 PivotTable1 所有字段的筛选器',
    '创建切片器': '为 PivotTable1 创建一个 Category 字段的切片器',
    '使用切片器筛选': '使用类别切片器筛选"电子产品"、"家居用品"、"服装"',
    '切换透视表布局': '将 PivotTable1 的布局切换为大纲格式',
    '获取透视表数据': '获取 PivotTable1 的数据区域和值',
    '格式化透视表': '格式化 PivotTable1，设置空单元格显示"--"，数据右对齐',
    '刷新透视表': '刷新 PivotTable1 的数据',
    '删除透视表': '删除 PivotTable1',
    
    // Excel 图表
    '创建柱状图': '基于 A1:B10 数据创建柱状图',
    '创建折线图': '基于 A1:B10 数据创建折线图',
    '柱状图': '基于 A1:B10 数据创建标题为"数据分析图表"的柱状图',
    '折线图': '基于 A1:B10 数据创建标题为"趋势分析"的折线图',
    
    // Excel 数据验证
    '设置数值验证': '限制 B2:B10 单元格只能输入大于 0 的整数',
    '设置下拉列表': '为 C2:C10 单元格设置下拉列表：选项1、选项2、选项3',
    
    // Excel 批注与命名
    '添加批注': '在 A1 单元格添加批注"这是一个批注。"',
    '创建命名区域': '将 A1:B10 区域命名为 MyData',
    
    // Word 文档操作
    '读取选中文本': '读取当前选中的文本内容',
    '读取整个文档': '读取整个文档的文本内容',
    '读取文档段落': '读取文档所有段落的内容',
    '在选中位置插入文本': '在当前选中的位置插入文本"Hello World"',
    '在文档末尾插入段落': '在文档末尾插入一个新段落"这是新的段落内容"',
    '插入多行文本（推荐）': '在文档末尾插入三行文本：第一行是"项目概述"，第二行是"这是一个示例项目"，第三行是"预计完成时间：2026年2月"',
    '在特定位置插入内容': '在文档开头插入"开头内容"，在末尾插入"末尾内容"',
    '插入 Base64 图片': '在文档末尾插入一张 400x300 的图片（使用 Base64 编码）',
    '创建列表': '创建一个包含"项目1"、"项目2"、"项目3"的项目符号列表',
    '插入表格': '创建一个 3 行 4 列的表格，包含标题和数据',
    '读取表格数据': '读取文档中第一个表格的所有数据',
    '创建内容控件': '将选中内容创建为标题为"客户名称"的内容控件',
    '读取/更新内容控件': '查找标签为 CustomerName 的内容控件，并将其内容更新为"Contoso Ltd."',
    '修改页眉': '在文档页眉中添加红色居中的文本"机密文件 - 仅供内部使用"',
    '设置文本格式': '将选中文本设置为微软雅黑、12号、加粗、深灰色',
    '设置段落格式': '设置所有段落为 1.5 倍行距、段后 10 磅间距、两端对齐',
    '简单替换': '把文档中所有的"旧版本"替换成"新版本"，不区分大小写',
    '高级搜索（通配符）': '搜索所有以 to 开头以 n 结尾的单词，并用黄色高亮显示',
    'word:创建表格': '创建一个 3 行 4 列的表格',
    'word:插入文本': '在光标位置插入文本"Hello World"',
    'word:插入段落': '在文档末尾插入新段落',
    'word:插入图片': '插入一张图片',
    
    // Word 字段操作
    '插入当前日期字段': '在光标位置插入当前日期字段，格式为"月/日/年"',
    '创建目录字段': '在文档开头插入目录',
    '插入超链接字段': '插入一个指向微软官网的超链接',
    '插入页码字段': '插入页码',
    'Addin字段（存储插件数据）': '创建一个用于存储插件自定义数据的 Addin 字段',
    '更新字段内容': '更新文档中所有日期字段',
    
    // Word 脚注和尾注
    '插入脚注': '在选中位置插入脚注"参考文献1"',
    '插入尾注': '在选中位置插入尾注"附录A"',
    '读取脚注内容': '搜索并统计文档中的脚注数量',
    
    // Word 样式管理
    '应用标题1样式': '将选中段落应用标题1样式',
    '应用标题样式（使用枚举）': '将选中段落应用标题2样式',
    '应用引用样式': '将选中段落应用引用块样式',
    '批量应用样式': '查找所有包含"重要"的段落，应用强调样式',
    '获取并应用现有样式': '获取第一个段落的样式，应用到选中区域',
    
    // Word 批注
    '插入语法批注': '为选中段落添加语法建议批注',
    '读取段落批注': '读取选中段落的所有批注',
    '注册批注事件': '监听批注被点击的事件',
    '删除批注': '删除选中段落的所有批注',
    '读取整个文档': '读取整个文档的文本内容',
    '读取文档段落': '读取文档所有段落的内容',
    '在选中位置插入文本': '在当前选中的位置插入文本"Hello World"',
    '在文档末尾插入段落': '在文档末尾插入一个新段落"这是新的段落内容"',
    '插入多行文本（推荐）': '在文档末尾插入三行文本：第一行是"项目概述"，第二行是"这是一个示例项目"，第三行是"预计完成时间：2026年2月"',
    '在特定位置插入内容': '在文档开头插入"开头内容"，在末尾插入"末尾内容"',
    '插入 Base64 图片': '在文档末尾插入一张 400x300 的图片（使用 Base64 编码）',
    '创建列表': '创建一个包含"项目1"、"项目2"、"项目3"的项目符号列表',
    '插入表格': '创建一个 3 行 4 列的表格，包含标题和数据',
    '读取表格数据': '读取文档中第一个表格的所有数据',
    '创建内容控件': '将选中内容创建为标题为"客户名称"的内容控件',
    '读取/更新内容控件': '查找标签为 CustomerName 的内容控件，并将其内容更新为"Contoso Ltd."',
    '修改页眉': '在文档页眉中添加红色居中的文本"机密文件 - 仅供内部使用"',
    'ppt:设置文本格式': '将选中文本设置为微软雅黑、12号、加粗、深灰色',
    'ppt:设置段落格式': '设置所有段落为 1.5 倍行距、段后 10 磅间距、两端对齐',
    'ppt:简单替换': '把文档中所有的"旧版本"替换成"新版本"，不区分大小写',
    'ppt:高级搜索（通配符）': '搜索所有以 to 开头以 n 结尾的单词，并用黄色高亮显示',
    'ppt:创建表格': '创建一个 3 行 4 列的表格',
    'ppt:插入文本': '在光标位置插入文本"Hello World"',
    'ppt:插入段落': '在文档末尾插入新段落',
    'ppt:插入图片': '插入一张图片',
    
    // PowerPoint 操作
    '添加幻灯片': '添加一张新的空白幻灯片',
    '删除幻灯片': '删除当前幻灯片',
    '添加文本框': '添加一个包含标题的文本框',
    '添加矩形': '添加一个蓝色矩形',
    
    // PowerPoint 文本格式化
    '设置文本粗体和斜体': '将第一个形状的文本设置为粗体和斜体',
    '设置字体名称和大小': '将第一个形状的字体设置为微软雅黑 24 号',
    '文本垂直居中': '将第一个形状的文本垂直居中对齐',
    '设置文本框边距': '设置第一个形状的文本框边距为上下5点、左右10点',
    '设置文本自动适应': '设置第一个形状的文本自动缩小以适应形状',
    
    // PowerPoint 布局和母版
    '获取幻灯片布局信息': '获取当前幻灯片的布局和母版名称',
    '使用指定布局创建幻灯片': '使用第一张幻灯片的布局创建新幻灯片',
    '获取所有可用布局': '列出所有可用的幻灯片布局',
    
    // PowerPoint 主题系统
    '使用主题颜色填充形状': '用主题强调色1填充第一个形状',
    '使用多种主题颜色': '创建3个矩形，分别使用主题强调色1、2、3',
    
    // PowerPoint 表格数据
    '读取表格数据': '读取第一个表格的所有单元格数据',
    '更新表格单元格': '将第一个表格的第一行第一列更新为"新内容"',
    '设置行高和列宽': '设置第一个表格的第一行高度为50点，第一列宽度为150点',
    '创建带合并单元格的表格': '创建一个3行4列的表格，第一行的4个单元格合并',
    '格式化表格单元格': '将第一个表格的标题行设置为蓝色背景、白色粗体文字',
  };
  
  const hostKey = `${hostType}:${name}`;
  return inputs[hostKey] || inputs[name] || `执行 ${name} 操作`;
}

/**
 * 生成期望行为
 */
function generateExpectedBehavior(name: string, category: string): string {
  const behaviors: Record<string, string> = {
    '创建工作表': '新工作表 NewSheet 被创建并激活',
    '重命名工作表': '当前工作表名称变为 RenamedSheet',
    '删除工作表': '工作表 SheetToDelete 被删除',
    '复制工作表': '当前工作表的副本被创建',
    
    '读取选中区域': '返回选中区域的值、地址和格式',
    '读取指定区域': '返回 A1:D10 区域的所有数据',
    '写入单个值': 'A1 单元格显示数字 100',
    
    'Excel 表格:创建表格': 'A1:D10 区域被转换为 Excel 表格',
    '创建柱状图': '工作表中出现一个柱状图',
    '创建透视表': 'F1 位置创建了数据透视表',
    
    '插入文本': '文档中出现指定文本',
    'Word 文档操作:创建表格': '文档中出现 3x4 表格',
    
    '添加幻灯片': '演示文稿中增加一张新幻灯片',
    '添加文本框': '幻灯片中出现文本框',
  };
  
  const key = `${category}:${name}`;
  return behaviors[key] || behaviors[name] || `${name} 操作成功完成`;
}

/**
 * 生成验证步骤
 */
function generateValidationSteps(name: string, category: string): string[] {
  return [
    '检查代码是否成功生成',
    '检查代码格式是否正确',
    '检查代码是否包含必要的 context.sync()',
    `验证 ${name} 的预期结果`,
  ];
}

/**
 * 确定测试用例优先级
 */
function determinePriority(category: string, name: string): 'high' | 'medium' | 'low' {
  const highPriorityCategories = ['工作表管理', '数据读取', '数据写入', '文档读取', '文本插入', '幻灯片操作'];
  const highPriorityNames = ['创建', '读取', '插入', '添加'];
  
  if (highPriorityCategories.includes(category)) {
    return 'high';
  }
  
  for (const keyword of highPriorityNames) {
    if (name.includes(keyword)) {
      return 'high';
    }
  }
  
  return 'medium';
}

/**
 * 生成测试套件
 */
function generateTestSuite(hostType: 'excel' | 'word' | 'powerpoint'): TestSuite {
  const projectRoot = path.resolve(__dirname, '../../..');
  const toolsFilePath = path.join(projectRoot, '.claude', 'skills', hostType, 'TOOLS.md');
  
  console.log(`读取文件: ${toolsFilePath}`);
  const content = fs.readFileSync(toolsFilePath, 'utf-8');
  
  const templates = parseToolsFile(content);
  console.log(`解析到 ${templates.length} 个代码模板`);
  
  const testCases = templates.map((template, index) => 
    generateTestCase(template, index, hostType)
  );
  
  return {
    id: `${hostType}-test-suite`,
    name: `${hostType.toUpperCase()} Test Suite`,
    hostType,
    testCases,
    metadata: {
      version: '1.0.0',
      generatedAt: new Date().toISOString(),
      sourceFile: toolsFilePath,
    },
  };
}

/**
 * 主函数
 */
function main() {
  console.log('开始生成测试用例...\n');
  
  const hostTypes: Array<'excel' | 'word' | 'powerpoint'> = ['excel', 'word', 'powerpoint'];
  
  for (const hostType of hostTypes) {
    try {
      console.log(`\n处理 ${hostType.toUpperCase()}...`);
      const testSuite = generateTestSuite(hostType);
      
      // 保存到 tests/office-skills/test-cases/ 目录（供脚本使用）
      const testCasesDir = path.resolve(__dirname, '../test-cases');
      if (!fs.existsSync(testCasesDir)) {
        fs.mkdirSync(testCasesDir, { recursive: true });
      }
      const testCasesPath = path.join(testCasesDir, `${hostType}-test-cases.json`);
      fs.writeFileSync(testCasesPath, JSON.stringify(testSuite, null, 2), 'utf-8');
      
      // 同时保存到 public/test-cases/ 目录（供网页访问）
      const publicDir = path.resolve(__dirname, '../../../public/test-cases');
      if (!fs.existsSync(publicDir)) {
        fs.mkdirSync(publicDir, { recursive: true });
      }
      const publicPath = path.join(publicDir, `${hostType}-test-cases.json`);
      fs.writeFileSync(publicPath, JSON.stringify(testSuite, null, 2), 'utf-8');
      
      console.log(`✅ 生成 ${testSuite.testCases.length} 个测试用例`);
      console.log(`   保存到: ${testCasesPath}`);
      console.log(`   网页访问: ${publicPath}`);
    } catch (error) {
      console.error(`❌ 处理 ${hostType} 失败:`, error);
    }
  }
  
  console.log('\n✅ 所有测试用例生成完成!');
}

// 运行
main();
