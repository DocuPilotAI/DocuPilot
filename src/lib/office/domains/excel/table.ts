/**
 * Excel Table 领域工具
 * 
 * 提供Excel表对象的创建、读取、编辑、排序、筛选等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { TableActionSchema } from "../../schemas/excel-schemas";
import * as actions from "../../codegen/excel/table-actions";
import { executeOfficeCode } from "../../executor";

export const excelTableTool = tool(
  "excel_table",
  `Excel表对象操作工具

支持的操作：
- create: 创建表格（将区域转换为表对象）
- read: 读取表格数据
- add_row: 添加行
- add_column: 添加列
- sort: 排序表格
- filter: 应用筛选
- delete: 删除表格

用于管理Excel的Table对象（不是普通单元格区域）。`,
  TableActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'create':
          code = actions.generateCreateCode(params);
          description = `创建表格 ${params.address}`;
          break;
        
        case 'read':
          code = actions.generateReadCode(params);
          description = `读取表格 "${params.name}"`;
          break;
        
        case 'add_row':
          code = actions.generateAddRowCode(params);
          description = `向表格 "${params.name}" 添加行`;
          break;
        
        case 'add_column':
          code = actions.generateAddColumnCode(params);
          description = `向表格 "${params.name}" 添加列 "${params.columnName}"`;
          break;
        
        case 'sort':
          code = actions.generateSortCode(params);
          description = `排序表格 "${params.name}" 的 "${params.column}" 列`;
          break;
        
        case 'filter':
          code = actions.generateFilterCode(params);
          description = `筛选表格 "${params.name}"`;
          break;
        
        case 'delete':
          code = actions.generateDeleteCode(params);
          description = `删除表格 "${params.name}"`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: create, read, add_row, add_column, sort, filter, delete`
            }]
          };
      }
      
      return executeOfficeCode('excel', code, description);
      
    } catch (error) {
      console.error('[excel_table] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
