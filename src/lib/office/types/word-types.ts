/**
 * Word 领域工具类型定义
 */

// ==================== Document 相关 ====================

export interface DocumentReadParams {
  includeFormat?: boolean;
  maxLength?: number;
}

export interface DocumentReadResult {
  text: string;
  paragraphCount: number;
  wordCount: number;
}

export interface DocumentSearchParams {
  searchText: string;
  matchCase?: boolean;
  matchWholeWord?: boolean;
}

export interface DocumentSearchResult {
  matches: Array<{
    text: string;
    context: string;
    paragraphIndex?: number;
  }>;
  count: number;
}

export interface DocumentReplaceParams {
  searchText: string;
  replaceText: string;
  matchCase?: boolean;
  replaceAll?: boolean;
}

export interface DocumentReplaceResult {
  replacedCount: number;
}

// ==================== Paragraph 相关 ====================

export type InsertLocation = 'Start' | 'End' | 'Before' | 'After' | 'Replace';
export type Alignment = 'Left' | 'Center' | 'Right' | 'Justify';

export interface ParagraphFormat {
  style?: string;  // "Heading 1", "Normal", etc.
  alignment?: Alignment;
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    color?: string;
  };
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
  };
}

export interface ParagraphInsertParams {
  text: string;
  location: InsertLocation;
  format?: ParagraphFormat;
}

export interface ParagraphInsertResult {
  success: boolean;
  paragraphIndex?: number;
}

// ==================== Table 相关 ====================

export interface TableCreateParams {
  rows: number;
  columns: number;
  data?: any[][];
  location?: InsertLocation;
  style?: string;
}

export interface TableCreateResult {
  tableIndex: number;
  success: boolean;
}

export interface TableReadResult {
  data: any[][];
  rows: number;
  columns: number;
}

export interface CellFormat {
  fill?: {
    color?: string;
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
