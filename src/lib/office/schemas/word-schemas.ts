/**
 * Word 领域工具 Zod Schema 定义
 */

import { z } from "zod";

// ==================== Document Schemas ====================

export const DocumentActionSchema = z.discriminatedUnion('action', [
  // read action
  z.object({
    action: z.literal('read'),
    includeFormat: z.boolean().optional().describe("是否包含格式信息"),
    maxLength: z.number().optional().describe("最大字符数")
  }),
  
  // read_selection action
  z.object({
    action: z.literal('read_selection')
  }),
  
  // search action
  z.object({
    action: z.literal('search'),
    searchText: z.string().describe("要搜索的文本"),
    matchCase: z.boolean().optional().describe("是否区分大小写"),
    matchWholeWord: z.boolean().optional().describe("是否匹配整个单词")
  }),
  
  // replace action
  z.object({
    action: z.literal('replace'),
    searchText: z.string(),
    replaceText: z.string(),
    matchCase: z.boolean().optional(),
    replaceAll: z.boolean().optional().describe("是否替换所有（默认false）")
  }),
  
  // clear action
  z.object({
    action: z.literal('clear')
  })
]);

// ==================== Paragraph Schemas ====================

const ParagraphFormatSchema = z.object({
  style: z.string().optional().describe("样式名称，如 'Heading 1', 'Normal'"),
  alignment: z.enum(['Left', 'Center', 'Right', 'Justify']).optional(),
  font: z.object({
    name: z.string().optional(),
    size: z.number().optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    color: z.string().optional()
  }).optional(),
  spacing: z.object({
    before: z.number().optional(),
    after: z.number().optional(),
    line: z.number().optional()
  }).optional()
});

export const ParagraphActionSchema = z.discriminatedUnion('action', [
  // insert action
  z.object({
    action: z.literal('insert'),
    text: z.string().describe("要插入的文本"),
    location: z.enum(['Start', 'End', 'Before', 'After', 'Replace']).describe("插入位置"),
    format: ParagraphFormatSchema.optional()
  }),
  
  // insert_at action
  z.object({
    action: z.literal('insert_at'),
    text: z.string(),
    index: z.number().describe("段落索引（从0开始）"),
    location: z.enum(['Before', 'After']),
    format: ParagraphFormatSchema.optional()
  }),
  
  // format action
  z.object({
    action: z.literal('format'),
    index: z.number().describe("段落索引"),
    format: ParagraphFormatSchema
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    index: z.number().describe("要删除的段落索引")
  }),
  
  // get action
  z.object({
    action: z.literal('get'),
    index: z.number().describe("段落索引")
  })
]);

// ==================== Table Schemas ====================

const CellFormatSchema = z.object({
  fill: z.object({
    color: z.string().optional()
  }).optional(),
  font: z.object({
    size: z.number().optional(),
    bold: z.boolean().optional(),
    color: z.string().optional()
  }).optional(),
  borders: z.object({
    color: z.string().optional(),
    weight: z.number().optional()
  }).optional()
});

export const TableActionSchema = z.discriminatedUnion('action', [
  // create action
  z.object({
    action: z.literal('create'),
    rows: z.number().describe("行数"),
    columns: z.number().describe("列数"),
    data: z.array(z.array(z.any())).optional().describe("表格数据（二维数组）"),
    location: z.enum(['Start', 'End']).optional(),
    style: z.string().optional().describe("表格样式名称")
  }),
  
  // read action
  z.object({
    action: z.literal('read'),
    tableIndex: z.number().describe("表格索引（从0开始）")
  }),
  
  // write action
  z.object({
    action: z.literal('write'),
    tableIndex: z.number(),
    data: z.array(z.array(z.any())).describe("要写入的数据")
  }),
  
  // insert_row action
  z.object({
    action: z.literal('insert_row'),
    tableIndex: z.number(),
    rowIndex: z.number(),
    values: z.array(z.any()).optional()
  }),
  
  // insert_column action
  z.object({
    action: z.literal('insert_column'),
    tableIndex: z.number(),
    columnIndex: z.number(),
    values: z.array(z.any()).optional()
  }),
  
  // delete_row action
  z.object({
    action: z.literal('delete_row'),
    tableIndex: z.number(),
    rowIndex: z.number()
  }),
  
  // delete_column action
  z.object({
    action: z.literal('delete_column'),
    tableIndex: z.number(),
    columnIndex: z.number()
  }),
  
  // format_cell action
  z.object({
    action: z.literal('format_cell'),
    tableIndex: z.number(),
    rowIndex: z.number(),
    columnIndex: z.number(),
    format: CellFormatSchema
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    tableIndex: z.number()
  })
]);
