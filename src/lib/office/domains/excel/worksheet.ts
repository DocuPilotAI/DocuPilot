/**
 * Excel Worksheet 领域工具
 * 
 * 提供工作表的增删改查、激活、复制等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { WorksheetActionSchema } from "../../schemas/excel-schemas";
import * as actions from "../../codegen/excel/worksheet-actions";
import { executeOfficeCode } from "../../executor";

export const excelWorksheetTool = tool(
  "excel_worksheet",
  `Excel工作表管理工具

支持的操作：
- list: 列出所有工作表
- add: 添加新工作表
- delete: 删除工作表
- rename: 重命名工作表
- exists: 检查工作表是否存在
- activate: 激活工作表
- copy: 复制工作表

用于管理Excel文件中的工作表（sheet/tab）。`,
  WorksheetActionSchema as any,  // 使用 as any 绕过类型检查
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'list':
          code = actions.generateListCode();
          description = '列出所有工作表';
          break;
        
        case 'add':
          code = actions.generateAddCode(params);
          description = `添加工作表 "${params.name}"`;
          break;
        
        case 'delete':
          code = actions.generateDeleteCode(params);
          description = `删除工作表 "${params.name}"`;
          break;
        
        case 'rename':
          code = actions.generateRenameCode(params);
          description = `重命名工作表 "${params.oldName}" 为 "${params.newName}"`;
          break;
        
        case 'exists':
          code = actions.generateExistsCode(params);
          description = `检查工作表 "${params.name}" 是否存在`;
          break;
        
        case 'activate':
          code = actions.generateActivateCode(params);
          description = `激活工作表 "${params.name}"`;
          break;
        
        case 'copy':
          code = actions.generateCopyCode(params);
          description = `复制工作表 "${params.sourceName}" 为 "${params.newName}"`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: list, add, delete, rename, exists, activate, copy`
            }]
          };
      }
      
      return executeOfficeCode('excel', code, description);
      
    } catch (error) {
      console.error('[excel_worksheet] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
