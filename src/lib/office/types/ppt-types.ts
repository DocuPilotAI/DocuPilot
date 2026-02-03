/**
 * PowerPoint 领域工具类型定义
 */

// ==================== Slide 相关 ====================

export type SlideLayout = 'Blank' | 'Title' | 'TitleAndContent';

export interface SlideInfo {
  index: number;
  id: string;
}

export interface SlideListResult {
  slideCount: number;
  slides: SlideInfo[];
}

export interface SlideReadResult {
  texts: Array<{
    content: string;
    shapeType: string;
  }>;
  shapes: Array<{
    id: string;
    type: string;
  }>;
  shapeCount: number;
}

export interface SlideAddParams {
  layout?: SlideLayout;
  insertAfter?: number;
}

export interface SlideAddResult {
  slideIndex: number;
  success: boolean;
}

// ==================== Shape 相关 ====================

export interface Position {
  left: number;
  top: number;
  width?: number;
  height?: number;
}

export interface TextFormat {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  fontName?: string;
  alignment?: 'Left' | 'Center' | 'Right';
  verticalAlignment?: 'Top' | 'Middle' | 'Bottom';
}

export interface ShapeFormat {
  fill?: {
    color: string;
  };
  line?: {
    color: string;
    weight: number;
    style: string;
  };
}

export type ShapeType = 'rectangle' | 'ellipse' | 'triangle' | 'rightArrow' | 'leftArrow' | 'upArrow' | 'downArrow' | 'star';

export interface ShapeAddTextParams {
  slideIndex: number;
  text: string;
  position: Position;
  format?: TextFormat;
}

export interface ShapeAddImageParams {
  slideIndex: number;
  base64Image: string;
  position: Position;
}

export interface ShapeAddParams {
  slideIndex: number;
  shapeType: ShapeType;
  position: Position;
  format?: ShapeFormat;
}

export interface ShapeResult {
  shapeId: string;
  success: boolean;
}

// ==================== Table 相关 ====================

export interface TableStyle {
  headerRow?: boolean;
  totalRow?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
}

export interface TableCreateParams {
  slideIndex: number;
  rows: number;
  columns: number;
  position: Position;
  data?: any[][];
  style?: TableStyle;
}

export interface TableCreateResult {
  tableId: string;
  success: boolean;
}

export interface TableCellFormat {
  fill?: {
    color: string;
  };
  font?: {
    size?: number;
    bold?: boolean;
    color?: string;
  };
  borders?: {
    color?: string;
    weight?: number;
  };
}
