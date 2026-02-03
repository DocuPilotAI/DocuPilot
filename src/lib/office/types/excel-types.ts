/**
 * Excel 领域工具类型定义
 */

// ==================== Range 相关 ====================

export interface RangeReadParams {
  address: string;
  includeFormulas?: boolean;
  includeFormat?: boolean;
}

export interface RangeReadResult {
  address: string;
  values: any[][];
  rowCount: number;
  columnCount: number;
  formulas?: any[][];
  numberFormat?: string[][];
}

export interface RangeWriteParams {
  address: string;
  values: any[][];
  autoExpand?: boolean;
}

export interface RangeWriteResult {
  success: boolean;
  writtenRange: string;
}

export interface RangeFormatParams {
  address: string;
  format: {
    font?: {
      name?: string;
      size?: number;
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
    };
    fill?: {
      color?: string;
    };
    borders?: {
      style?: string;
      color?: string;
      weight?: string;
    };
    numberFormat?: string;
    horizontalAlignment?: 'Left' | 'Center' | 'Right' | 'Justify';
    verticalAlignment?: 'Top' | 'Center' | 'Bottom';
  };
}

// ==================== Worksheet 相关 ====================

export interface WorksheetInfo {
  name: string;
  id: string;
  position: number;
  visible: boolean;
}

export interface WorksheetListResult {
  sheets: WorksheetInfo[];
  count: number;
}

export interface WorksheetAddParams {
  name: string;
  position?: 'start' | 'end' | 'before' | 'after';
  referenceSheet?: string;
}

export interface WorksheetAddResult {
  name: string;
  id: string;
  success: boolean;
}

// ==================== Table 相关 ====================

export interface TableCreateParams {
  address: string;
  hasHeaders: boolean;
  name?: string;
  style?: string;
}

export interface TableReadResult {
  name: string;
  headers: string[];
  data: any[][];
  totalRows: number;
}

// ==================== Chart 相关 ====================

export type ChartType = 'columnClustered' | 'columnStacked' | 'line' | 'lineMarkers' | 'pie' | 'barClustered' | 'barStacked' | 'area' | 'areaStacked' | 'scatter';

export interface ChartCreateParams {
  dataRange: string;
  chartType: ChartType;
  position?: {
    left: number;
    top: number;
  };
  title?: string;
}

export interface ChartCreateResult {
  chartId: string;
  success: boolean;
}
