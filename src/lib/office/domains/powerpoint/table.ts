/**
 * PowerPoint Table 领域工具
 * 
 * 提供PowerPoint表格的创建、编辑、格式化等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { TableActionSchema } from "../../schemas/ppt-schemas";
import * as actions from "../../codegen/powerpoint/table-actions";
import { executeOfficeCode } from "../../executor";

export const pptTableTool = tool(
  "ppt_table",
  `PowerPoint表格操作工具

支持的操作：
- create: 创建表格
- read: 读取表格数据
- write: 写入表格数据
- format_cell: 格式化单元格
- delete: 删除表格

用于在PowerPoint幻灯片中创建和管理表格。`,
  TableActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'create':
          code = actions.generateCreateCode(params);
          description = `在幻灯片 ${params.slideIndex} 创建 ${params.rows}x${params.columns} 表格`;
          break;
        
        case 'read':
          code = actions.generateReadCode(params);
          description = `读取幻灯片 ${params.slideIndex} 的表格`;
          break;
        
        case 'write':
          code = actions.generateWriteCode(params);
          description = `写入表格数据`;
          break;
        
        case 'format_cell':
          code = actions.generateFormatCellCode(params);
          description = `格式化表格单元格 [${params.row},${params.column}]`;
          break;
        
        case 'delete':
          code = actions.generateDeleteTableCode(params);
          description = `删除表格`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: create, read, write, format_cell, delete`
            }]
          };
      }
      
      return executeOfficeCode('powerpoint', code, description);
      
    } catch (error) {
      console.error('[ppt_table] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
