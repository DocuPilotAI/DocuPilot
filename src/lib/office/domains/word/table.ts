/**
 * Word Table 领域工具
 * 
 * 提供Word表格的创建、编辑、格式化等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { TableActionSchema } from "../../schemas/word-schemas";
import * as actions from "../../codegen/word/table-actions";
import { executeOfficeCode } from "../../executor";

export const wordTableTool = tool(
  "word_table",
  `Word表格操作工具

支持的操作：
- create: 创建表格
- read: 读取表格数据
- write: 写入表格数据
- insert_row: 插入行
- insert_column: 插入列
- delete_row: 删除行
- delete_column: 删除列
- format_cell: 格式化单元格
- delete: 删除表格

用于在Word文档中创建和管理表格。`,
  TableActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'create':
          code = actions.generateCreateCode(params);
          description = `创建 ${params.rows}x${params.columns} 表格`;
          break;
        
        case 'read':
          code = actions.generateReadCode(params);
          description = `读取表格 ${params.tableIndex}`;
          break;
        
        case 'write':
          code = actions.generateWriteCode(params);
          description = `写入表格 ${params.tableIndex} 数据`;
          break;
        
        case 'insert_row':
          code = actions.generateInsertRowCode(params);
          description = `向表格 ${params.tableIndex} 插入行`;
          break;
        
        case 'insert_column':
          code = actions.generateInsertColumnCode(params);
          description = `向表格 ${params.tableIndex} 插入列`;
          break;
        
        case 'delete_row':
          code = actions.generateDeleteRowCode(params);
          description = `删除表格 ${params.tableIndex} 的行 ${params.rowIndex}`;
          break;
        
        case 'delete_column':
          code = actions.generateDeleteColumnCode(params);
          description = `删除表格 ${params.tableIndex} 的列 ${params.columnIndex}`;
          break;
        
        case 'format_cell':
          code = actions.generateFormatCellCode(params);
          description = `格式化表格 ${params.tableIndex} 单元格 [${params.rowIndex},${params.columnIndex}]`;
          break;
        
        case 'delete':
          code = actions.generateDeleteTableCode(params);
          description = `删除表格 ${params.tableIndex}`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: create, read, write, insert_row, insert_column, delete_row, delete_column, format_cell, delete`
            }]
          };
      }
      
      return executeOfficeCode('word', code, description);
      
    } catch (error) {
      console.error('[word_table] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
