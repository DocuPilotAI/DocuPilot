/**
 * PowerPoint 领域工具 Zod Schema 定义
 */

import { z } from "zod";

// ==================== Slide Schemas ====================

export const SlideActionSchema = z.discriminatedUnion('action', [
  // list action
  z.object({
    action: z.literal('list')
  }),
  
  // read action
  z.object({
    action: z.literal('read'),
    slideIndex: z.number().int().min(0).describe("幻灯片索引（从0开始）")
  }),
  
  // add action
  z.object({
    action: z.literal('add'),
    layout: z.enum(['Blank', 'Title', 'TitleAndContent']).optional().describe("幻灯片布局"),
    insertAfter: z.number().optional().describe("在指定索引后插入")
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    slideIndex: z.number().int().min(0)
  }),
  
  // duplicate action
  z.object({
    action: z.literal('duplicate'),
    slideIndex: z.number().int().min(0)
  }),
  
  // move action
  z.object({
    action: z.literal('move'),
    fromIndex: z.number().int().min(0),
    toIndex: z.number().int().min(0)
  })
]);

// ==================== Shape Schemas ====================

const PositionSchema = z.object({
  left: z.number().describe("左边距（points）"),
  top: z.number().describe("上边距（points）"),
  width: z.number().optional().describe("宽度（points）"),
  height: z.number().optional().describe("高度（points）")
});

const TextFormatSchema = z.object({
  fontSize: z.number().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string().optional(),
  fontName: z.string().optional(),
  alignment: z.enum(['Left', 'Center', 'Right']).optional(),
  verticalAlignment: z.enum(['Top', 'Middle', 'Bottom']).optional()
});

const ShapeFormatSchema = z.object({
  fill: z.object({
    color: z.string()
  }).optional(),
  line: z.object({
    color: z.string(),
    weight: z.number(),
    style: z.string()
  }).optional()
});

export const ShapeActionSchema = z.discriminatedUnion('action', [
  // add_text action
  z.object({
    action: z.literal('add_text'),
    slideIndex: z.number().int().min(0),
    text: z.string().describe("文本内容"),
    position: PositionSchema,
    format: TextFormatSchema.optional()
  }),
  
  // add_image action
  z.object({
    action: z.literal('add_image'),
    slideIndex: z.number().int().min(0),
    base64Image: z.string().describe("Base64编码的图片"),
    position: PositionSchema
  }),
  
  // add_shape action
  z.object({
    action: z.literal('add_shape'),
    slideIndex: z.number().int().min(0),
    shapeType: z.enum(['rectangle', 'ellipse', 'triangle', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'star']),
    position: PositionSchema,
    format: ShapeFormatSchema.optional()
  }),
  
  // update_text action
  z.object({
    action: z.literal('update_text'),
    slideIndex: z.number().int().min(0),
    shapeId: z.string(),
    text: z.string()
  }),
  
  // format action
  z.object({
    action: z.literal('format'),
    slideIndex: z.number().int().min(0),
    shapeId: z.string(),
    format: ShapeFormatSchema
  }),
  
  // move action
  z.object({
    action: z.literal('move'),
    slideIndex: z.number().int().min(0),
    shapeId: z.string(),
    position: PositionSchema
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    slideIndex: z.number().int().min(0),
    shapeId: z.string()
  })
]);

// ==================== Table Schemas ====================

const TableStyleSchema = z.object({
  headerRow: z.boolean().optional(),
  totalRow: z.boolean().optional(),
  firstColumn: z.boolean().optional(),
  lastColumn: z.boolean().optional(),
  bandedRows: z.boolean().optional(),
  bandedColumns: z.boolean().optional()
});

const TableCellFormatSchema = z.object({
  fill: z.object({
    color: z.string()
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
    slideIndex: z.number().int().min(0),
    rows: z.number().int().min(1),
    columns: z.number().int().min(1),
    position: PositionSchema,
    data: z.array(z.array(z.any())).optional(),
    style: TableStyleSchema.optional()
  }),
  
  // read action
  z.object({
    action: z.literal('read'),
    slideIndex: z.number().int().min(0),
    tableId: z.string()
  }),
  
  // write action
  z.object({
    action: z.literal('write'),
    slideIndex: z.number().int().min(0),
    tableId: z.string(),
    data: z.array(z.array(z.any()))
  }),
  
  // format_cell action
  z.object({
    action: z.literal('format_cell'),
    slideIndex: z.number().int().min(0),
    tableId: z.string(),
    row: z.number().int().min(0),
    column: z.number().int().min(0),
    format: TableCellFormatSchema
  }),
  
  // delete action
  z.object({
    action: z.literal('delete'),
    slideIndex: z.number().int().min(0),
    tableId: z.string()
  })
]);
