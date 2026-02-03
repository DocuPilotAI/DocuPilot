import { OfficeHostType } from "./host-detector";

// Excel 专用工具
export const excelTools = [
  { name: "office_excel_read_range", description: "读取 Excel 单元格区域" },
  { name: "office_excel_write_range", description: "写入 Excel 单元格区域" },
  { name: "office_excel_get_selection", description: "获取当前选中区域" },
  { name: "office_excel_create_chart", description: "创建图表" },
  { name: "office_excel_add_formula", description: "添加公式" },
  { name: "office_excel_format_cells", description: "格式化单元格" },
];

// Word 专用工具
export const wordTools = [
  { name: "office_word_read_document", description: "读取 Word 文档内容" },
  { name: "office_word_insert_text", description: "插入文本" },
  { name: "office_word_insert_paragraph", description: "插入段落" },
  { name: "office_word_insert_table", description: "插入表格" },
  { name: "office_word_format_text", description: "格式化文本" },
  { name: "office_word_search_replace", description: "搜索替换" },
];

// PowerPoint 专用工具
export const powerpointTools = [
  { name: "office_ppt_get_slides", description: "获取幻灯片列表" },
  { name: "office_ppt_add_slide", description: "添加幻灯片" },
  { name: "office_ppt_delete_slide", description: "删除幻灯片" },
  { name: "office_ppt_add_text", description: "添加文本框" },
  { name: "office_ppt_add_shape", description: "添加形状" },
  { name: "office_ppt_add_image", description: "添加图片" },
];

// 根据宿主类型获取可用工具列表
export function getToolsForHost(hostType: OfficeHostType): string[] {
  const commonTools = [
    "execute_python_code",
    "create_chart",
    "describe_data",
  ];
  
  const hostTools: Record<string, string[]> = {
    excel: excelTools.map(t => t.name),
    word: wordTools.map(t => t.name),
    powerpoint: powerpointTools.map(t => t.name),
  };
  
  return [...commonTools, ...(hostTools[hostType] || [])];
}

// 获取工具描述
export function getToolDescription(toolName: string): string {
  const allTools = [...excelTools, ...wordTools, ...powerpointTools];
  const tool = allTools.find(t => t.name === toolName);
  return tool?.description || toolName;
}
