/**
 * Excel 领域工具 Zod Schema 定义
 */

import { z } from "zod";

// ==================== Range Schemas ====================

const RangeAddressSchema = z.string()
  .regex(/^[A-Z]+\d+(:[A-Z]+\d+)?$/, "必须是有效的A1表示法，如 'A1' 或 'A1:B10'")
  .describe("单元格地址");

const RangeFormatSchema = z.object({
  font: z.object({
    name: z.string().optional(),
    size: z.number().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    color: z.string().optional()
  }).optional(),
  fill: z.object({
    color: z.string().optional()
  }).optional(),
  borders: z.object({
    style: z.string().optional(),
    color: z.string().optional(),
    weight: z.string().optional()
  }).optional(),
  numberFormat: z.string().optional(),
  horizontalAlignment: z.enum(['Left', 'Center', 'Right', 'Justify']).optional(),
  verticalAlignment: z.enum(['Top', 'Center', 'Bottom']).optional()
});

export const RangeActionSchema = z.discriminatedUnion('action', [
  // read action
  z.object({
    action: z.literal('read'),
    address: RangeAddressSchema,
    includeFormulas: z.boolean().optional().describe("是否包含公式"),
    includeFormat: z.boolean().optional().describe("是否包含格式")
  }),
  
  // read_selection action
  z.object({
    action: z.literal('read_selection')
  }),
  
  // write action
  z.object({
    action: z.literal('write'),
    address: RangeAddressSchema,
    values: z.array(z.array(z.any())).describe("二维数组数据"),
    autoExpand: z.boolean().optional().describe("是否自动扩展区域（默认true）")
  }),
  
  // format action
  z.object({
    action: z.literal('format'),
    address: RangeAddressSchema,
    format: RangeFormatSchema
  }),
  
  // clear action
  z.object({
    action: z.literal('clear'),
    address: RangeAddressSchema,
    applyTo: z.enum(['contents', 'formats', 'all']).optional().describe("清除内容/格式/全部")
  }),
  
  // copy action
  z.object({
    action: z.literal('copy'),
    source: RangeAddressSchema,
    destination: RangeAddressSchema,
    copyType: z.enum(['all', 'values', 'formats']).optional().describe("复制类型")
  }),
  
  // insert action
  z.object({
    action: z.literal('insert'),
    address: RangeAddressSchema,
    shift: z.enum(['down', 'right']).describe("插入后单元格移动方向")
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    address: RangeAddressSchema,
    shift: z.enum(['up', 'left']).describe("删除后单元格移动方向")
  })
]);

// ==================== Worksheet Schemas ====================

export const WorksheetActionSchema = z.discriminatedUnion('action', [
  // list action
  z.object({
    action: z.literal('list')
  }),
  
  // add action
  z.object({
    action: z.literal('add'),
    name: z.string().describe("工作表名称"),
    position: z.enum(['start', 'end', 'before', 'after']).optional(),
    referenceSheet: z.string().optional().describe("参考工作表名称（用于before/after）")
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    name: z.string().describe("要删除的工作表名称")
  }),
  
  // rename action
  z.object({
    action: z.literal('rename'),
    oldName: z.string(),
    newName: z.string()
  }),
  
  // exists action
  z.object({
    action: z.literal('exists'),
    name: z.string().describe("工作表名称")
  }),
  
  // activate action
  z.object({
    action: z.literal('activate'),
    name: z.string().describe("要激活的工作表名称")
  }),
  
  // copy action
  z.object({
    action: z.literal('copy'),
    sourceName: z.string(),
    newName: z.string(),
    position: z.enum(['before', 'after']).optional()
  })
]);

// ==================== Table Schemas ====================

export const TableActionSchema = z.discriminatedUnion('action', [
  // create action
  z.object({
    action: z.literal('create'),
    address: RangeAddressSchema,
    hasHeaders: z.boolean().describe("是否包含标题行"),
    name: z.string().optional().describe("表格名称"),
    style: z.string().optional().describe("表格样式")
  }),
  
  // read action
  z.object({
    action: z.literal('read'),
    name: z.string().describe("表格名称"),
    includeHeaders: z.boolean().optional().describe("是否包含标题行")
  }),
  
  // add_row action
  z.object({
    action: z.literal('add_row'),
    name: z.string(),
    values: z.array(z.any()).optional(),
    index: z.number().optional().describe("插入位置索引")
  }),
  
  // add_column action
  z.object({
    action: z.literal('add_column'),
    name: z.string(),
    columnName: z.string().describe("列名称"),
    values: z.array(z.any()).optional()
  }),
  
  // sort action
  z.object({
    action: z.literal('sort'),
    name: z.string(),
    column: z.string().describe("排序列名"),
    ascending: z.boolean().describe("是否升序")
  }),
  
  // filter action
  z.object({
    action: z.literal('filter'),
    name: z.string(),
    column: z.string(),
    criteria: z.any().describe("筛选条件")
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    name: z.string().describe("表格名称")
  })
]);

// ==================== Chart Schemas ====================

export const ChartActionSchema = z.discriminatedUnion('action', [
  // create action
  z.object({
    action: z.literal('create'),
    dataRange: RangeAddressSchema,
    chartType: z.enum([
      'columnClustered', 'columnStacked', 'line', 'lineMarkers', 
      'pie', 'barClustered', 'barStacked', 'area', 'areaStacked', 'scatter'
    ]).describe("图表类型"),
    position: z.object({
      left: z.number(),
      top: z.number()
    }).optional(),
    title: z.string().optional()
  }),
  
  // update action
  z.object({
    action: z.literal('update'),
    chartId: z.string(),
    dataRange: RangeAddressSchema.optional()
  }),
  
  // set_title action
  z.object({
    action: z.literal('set_title'),
    chartId: z.string(),
    title: z.string()
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    chartId: z.string()
  })
]);
