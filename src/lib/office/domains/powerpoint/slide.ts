/**
 * PowerPoint Slide 领域工具
 * 
 * 提供幻灯片管理：列表、读取、添加、删除、复制、移动
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { SlideActionSchema } from "../../schemas/ppt-schemas";
import * as actions from "../../codegen/powerpoint/slide-actions";
import { executeOfficeCode } from "../../executor";

export const pptSlideTool = tool(
  "ppt_slide",
  `PowerPoint幻灯片管理工具

支持的操作：
- list: 列出所有幻灯片
- read: 读取指定幻灯片的内容
- add: 添加新幻灯片
- delete: 删除幻灯片
- duplicate: 复制幻灯片
- move: 移动幻灯片位置

用于管理PowerPoint演示文稿中的幻灯片。`,
  SlideActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'list':
          code = actions.generateListCode();
          description = '列出所有幻灯片';
          break;
        
        case 'read':
          code = actions.generateReadCode(params);
          description = `读取幻灯片 ${params.slideIndex} 的内容`;
          break;
        
        case 'add':
          code = actions.generateAddCode(params);
          description = `添加新幻灯片`;
          break;
        
        case 'delete':
          code = actions.generateDeleteCode(params);
          description = `删除幻灯片 ${params.slideIndex}`;
          break;
        
        case 'duplicate':
          code = actions.generateDuplicateCode(params);
          description = `复制幻灯片 ${params.slideIndex}`;
          break;
        
        case 'move':
          code = actions.generateMoveCode(params);
          description = `移动幻灯片 ${params.fromIndex} 到 ${params.toIndex}`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: list, read, add, delete, duplicate, move`
            }]
          };
      }
      
      return executeOfficeCode('powerpoint', code, description);
      
    } catch (error) {
      console.error('[ppt_slide] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
