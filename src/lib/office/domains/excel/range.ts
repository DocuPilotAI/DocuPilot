/**
 * Excel Range 领域工具
 * 
 * 提供单元格区域的读写、格式化、复制、插入删除等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { RangeActionSchema } from "../../schemas/excel-schemas";
import * as actions from "../../codegen/excel/range-actions";
import { executeOfficeCode } from "../../executor";

export const excelRangeTool = tool(
  "excel_range",
  `Excel单元格区域操作工具

支持的操作：
- read: 读取指定区域数据
- read_selection: 读取当前选中区域
- write: 写入数据到区域
- format: 格式化单元格（字体、填充、边框、数字格式等）
- clear: 清除内容/格式
- copy: 复制区域
- insert: 插入单元格
- delete: 删除单元格

这是最高频使用的Excel工具，覆盖大部分数据操作需求。`,
  RangeActionSchema as any,  // 使用 as any 绕过类型检查
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'read':
          code = actions.generateReadCode(params);
          description = `读取区域 ${params.address}`;
          break;
        
        case 'read_selection':
          code = actions.generateReadSelectionCode();
          description = '读取选中区域';
          break;
        
        case 'write':
          code = actions.generateWriteCode(params);
          description = `写入数据到 ${params.address}`;
          break;
        
        case 'format':
          code = actions.generateFormatCode(params);
          description = `格式化区域 ${params.address}`;
          break;
        
        case 'clear':
          code = actions.generateClearCode(params);
          description = `清除 ${params.address} 的${params.applyTo || '所有内容'}`;
          break;
        
        case 'copy':
          code = actions.generateCopyCode(params);
          description = `复制 ${params.source} 到 ${params.destination}`;
          break;
        
        case 'insert':
          code = actions.generateInsertDeleteCode('insert', params);
          description = `在 ${params.address} 插入单元格`;
          break;
        
        case 'delete':
          code = actions.generateInsertDeleteCode('delete', params);
          description = `删除 ${params.address} 的单元格`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: read, read_selection, write, format, clear, copy, insert, delete`
            }]
          };
      }
      
      return executeOfficeCode('excel', code, description);
      
    } catch (error) {
      console.error('[excel_range] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
