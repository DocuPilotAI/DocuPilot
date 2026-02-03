/**
 * Excel Chart 领域工具
 * 
 * 提供图表的创建、更新、删除等操作
 */

import { tool } from "@anthropic-ai/claude-agent-sdk";
import { ChartActionSchema } from "../../schemas/excel-schemas";
import * as actions from "../../codegen/excel/chart-actions";
import { executeOfficeCode } from "../../executor";

export const excelChartTool = tool(
  "excel_chart",
  `Excel图表操作工具

支持的操作：
- create: 创建图表（柱状图、折线图、饼图等）
- update: 更新图表数据源
- set_title: 设置图表标题
- delete: 删除图表

用于在Excel中创建和管理各种类型的图表。`,
  ChartActionSchema as any,
  async (args) => {
    const { action, ...params } = args as any;
    
    let code: string;
    let description: string;
    
    try {
      switch (action) {
        case 'create':
          code = actions.generateCreateCode(params);
          description = `创建 ${params.chartType} 图表`;
          break;
        
        case 'update':
          code = actions.generateUpdateCode(params);
          description = `更新图表 ${params.chartId}`;
          break;
        
        case 'set_title':
          code = actions.generateSetTitleCode(params);
          description = `设置图表标题 "${params.title}"`;
          break;
        
        case 'delete':
          code = actions.generateDeleteCode(params);
          description = `删除图表 ${params.chartId}`;
          break;
        
        default:
          return {
            content: [{
              type: "text" as const,
              text: `❌ 不支持的操作: ${action}\n\n支持的操作: create, update, set_title, delete`
            }]
          };
      }
      
      return executeOfficeCode('excel', code, description);
      
    } catch (error) {
      console.error('[excel_chart] Error generating code:', error);
      return {
        content: [{
          type: "text" as const,
          text: `❌ 生成代码时出错: ${error instanceof Error ? error.message : String(error)}`
        }]
      };
    }
  }
);
