/**
 * PowerPoint Shape 领域工具
 * 
 * 提供形状和文本框操作：添加文本、图片、形状，更新、格式化、移动、删除
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { ShapeActionSchema } from "../../schemas/ppt-schemas";
import * as actions from "../../codegen/powerpoint/shape-actions";
import { executeOfficeCode } from "../../executor";

export const pptShapeTool = tool(
  "ppt_shape",
  `PowerPoint形状和文本框操作工具

支持的操作：
- add_text: 添加文本框
- add_image: 添加图片
- add_shape: 添加几何形状（矩形、圆形、箭头等）
- update_text: 更新文本内容
- format: 格式化形状
- move: 移动形状位置
- delete: 删除形状

这是PowerPoint中最高频的工具，用于向幻灯片添加和管理内容。`,
  ShapeActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'add_text':
          code = actions.generateAddTextCode(params);
          description = `在幻灯片 ${params.slideIndex} 添加文本框`;
          break;
        
        case 'add_image':
          code = actions.generateAddImageCode(params);
          description = `在幻灯片 ${params.slideIndex} 添加图片`;
          break;
        
        case 'add_shape':
          code = actions.generateAddShapeCode(params);
          description = `在幻灯片 ${params.slideIndex} 添加 ${params.shapeType} 形状`;
          break;
        
        case 'update_text':
          code = actions.generateUpdateTextCode(params);
          description = `更新幻灯片 ${params.slideIndex} 的文本`;
          break;
        
        case 'format':
          code = actions.generateFormatShapeCode(params);
          description = `格式化幻灯片 ${params.slideIndex} 的形状`;
          break;
        
        case 'move':
          code = actions.generateMoveShapeCode(params);
          description = `移动幻灯片 ${params.slideIndex} 的形状`;
          break;
        
        case 'delete':
          code = actions.generateDeleteShapeCode(params);
          description = `删除幻灯片 ${params.slideIndex} 的形状`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: add_text, add_image, add_shape, update_text, format, move, delete`
            }]
          };
      }
      
      return executeOfficeCode('powerpoint', code, description);
      
    } catch (error) {
      console.error('[ppt_shape] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
