/**
 * Word Paragraph 领域工具
 * 
 * 提供段落操作：插入、格式化、删除
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { ParagraphActionSchema } from "../../schemas/word-schemas";
import * as actions from "../../codegen/word/paragraph-actions";
import { executeOfficeCode } from "../../executor";

export const wordParagraphTool = tool(
  "word_paragraph",
  `Word段落操作工具

支持的操作：
- insert: 插入段落到指定位置（Start/End/Before/After/Replace）
- insert_at: 在指定段落索引位置插入
- format: 格式化指定段落
- delete: 删除指定段落
- get: 获取段落信息

这是Word中最高频的工具，用于处理段落和文本内容。`,
  ParagraphActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'insert':
          code = actions.generateInsertCode(params);
          description = `插入段落到 ${params.location}`;
          break;
        
        case 'insert_at':
          code = actions.generateInsertAtCode(params);
          description = `在段落 ${params.index} ${params.location === 'Before' ? '之前' : '之后'} 插入`;
          break;
        
        case 'format':
          code = actions.generateFormatParagraphCode(params);
          description = `格式化段落 ${params.index}`;
          break;
        
        case 'delete':
          code = actions.generateDeleteCode(params);
          description = `删除段落 ${params.index}`;
          break;
        
        case 'get':
          code = actions.generateGetCode(params);
          description = `获取段落 ${params.index} 信息`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: insert, insert_at, format, delete, get`
            }]
          };
      }
      
      return executeOfficeCode('word', code, description);
      
    } catch (error) {
      console.error('[word_paragraph] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
