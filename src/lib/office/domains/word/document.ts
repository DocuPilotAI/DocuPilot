/**
 * Word Document 领域工具
 * 
 * 提供文档级操作：读取、搜索、替换、清除
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { DocumentActionSchema } from "../../schemas/word-schemas";
import * as actions from "../../codegen/word/document-actions";
import { executeOfficeCode } from "../../executor";

export const wordDocumentTool = tool(
  "word_document",
  `Word文档级操作工具

支持的操作：
- read: 读取文档内容
- read_selection: 读取当前选中文本
- search: 搜索文本
- replace: 替换文本
- clear: 清除文档内容

用于读取和管理整个Word文档的内容。`,
  DocumentActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'read':
          code = actions.generateReadCode(params);
          description = '读取文档内容';
          break;
        
        case 'read_selection':
          code = actions.generateReadSelectionCode();
          description = '读取选中文本';
          break;
        
        case 'search':
          code = actions.generateSearchCode(params);
          description = `搜索文本 "${params.searchText}"`;
          break;
        
        case 'replace':
          code = actions.generateReplaceCode(params);
          description = `替换 "${params.searchText}" 为 "${params.replaceText}"`;
          break;
        
        case 'clear':
          code = actions.generateClearCode();
          description = '清除文档内容';
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: read, read_selection, search, replace, clear`
            }]
          };
      }
      
      return executeOfficeCode('word', code, description);
      
    } catch (error) {
      console.error('[word_document] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
