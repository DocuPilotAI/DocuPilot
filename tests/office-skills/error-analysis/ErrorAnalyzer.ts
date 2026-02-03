/**
 * 错误分析器 - 分析错误模式并生成优化建议
 */

import type { ErrorReport, ErrorStatistics } from '../test-runner/types';
import { ErrorCollector } from './ErrorCollector';

export interface AnalysisResult {
  summary: {
    totalErrors: number;
    errorRate: number;
    uniqueTestCases: number;
    timeRange: {
      start: string;
      end: string;
    };
  };
  statistics: ErrorStatistics;
  patterns: ErrorPattern[];
  recommendations: Recommendation[];
  topProblematicAreas: ProblematicArea[];
}

export interface ErrorPattern {
  pattern: string;
  description: string;
  frequency: number;
  affectedTestCases: string[];
  examples: string[];
}

export interface Recommendation {
  priority: 'high' | 'medium' | 'low';
  area: string;
  issue: string;
  suggestion: string;
  affectedCount: number;
}

export interface ProblematicArea {
  category: string;
  errorCount: number;
  errorRate: number;
  topErrors: Array<{
    message: string;
    count: number;
  }>;
}

export class ErrorAnalyzer {
  /**
   * 分析所有错误报告
   */
  static analyzeAll(): AnalysisResult {
    const reports = ErrorCollector.getAllReports();
    
    return {
      summary: this.generateSummary(reports),
      statistics: this.generateStatistics(reports),
      patterns: this.identifyPatterns(reports),
      recommendations: this.generateRecommendations(reports),
      topProblematicAreas: this.identifyProblematicAreas(reports),
    };
  }

  /**
   * 按主机类型分析
   */
  static analyzeByHost(hostType: 'excel' | 'word' | 'powerpoint'): AnalysisResult {
    const reports = ErrorCollector.getReportsByHost(hostType);
    
    return {
      summary: this.generateSummary(reports),
      statistics: this.generateStatistics(reports),
      patterns: this.identifyPatterns(reports),
      recommendations: this.generateRecommendations(reports),
      topProblematicAreas: this.identifyProblematicAreas(reports),
    };
  }

  /**
   * 生成摘要
   */
  private static generateSummary(reports: ErrorReport[]) {
    if (reports.length === 0) {
      return {
        totalErrors: 0,
        errorRate: 0,
        uniqueTestCases: 0,
        timeRange: { start: '', end: '' },
      };
    }

    const timestamps = reports.map(r => new Date(r.timestamp).getTime());
    const uniqueTestCases = new Set(reports.map(r => r.testCaseId).filter(Boolean)).size;

    return {
      totalErrors: reports.length,
      errorRate: 1.0, // 这里需要知道总测试数才能计算准确的错误率
      uniqueTestCases,
      timeRange: {
        start: new Date(Math.min(...timestamps)).toISOString(),
        end: new Date(Math.max(...timestamps)).toISOString(),
      },
    };
  }

  /**
   * 生成统计信息
   */
  private static generateStatistics(reports: ErrorReport[]): ErrorStatistics {
    const errorsByType: Record<string, number> = {};
    const errorsByCategory: Record<string, number> = {};
    const errorMessages = new Map<string, { count: number; testCases: Set<string> }>();

    for (const report of reports) {
      // 按类型统计
      errorsByType[report.errorType] = (errorsByType[report.errorType] || 0) + 1;

      // 按类别统计（从 testCaseId 提取）
      if (report.testCaseId) {
        const category = this.extractCategory(report.testCaseId);
        errorsByCategory[category] = (errorsByCategory[category] || 0) + 1;
      }

      // 统计错误消息
      if (!errorMessages.has(report.errorMessage)) {
        errorMessages.set(report.errorMessage, {
          count: 0,
          testCases: new Set(),
        });
      }
      const msgData = errorMessages.get(report.errorMessage)!;
      msgData.count++;
      if (report.testCaseId) {
        msgData.testCases.add(report.testCaseId);
      }
    }

    // 获取 Top 10 错误
    const topErrors = Array.from(errorMessages.entries())
      .map(([message, data]) => ({
        message,
        count: data.count,
        testCases: Array.from(data.testCases),
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    return {
      totalErrors: reports.length,
      errorsByType: errorsByType as any,
      errorsByCategory,
      topErrors,
    };
  }

  /**
   * 识别错误模式
   */
  private static identifyPatterns(reports: ErrorReport[]): ErrorPattern[] {
    const patterns: ErrorPattern[] = [];

    // 模式 1: InvalidArgument 错误
    const invalidArgErrors = reports.filter(r => r.errorType === 'InvalidArgument');
    if (invalidArgErrors.length > 0) {
      patterns.push({
        pattern: 'InvalidArgument',
        description: '参数无效或缺少错误',
        frequency: invalidArgErrors.length,
        affectedTestCases: Array.from(new Set(invalidArgErrors.map(r => r.testCaseId).filter(Boolean))) as string[],
        examples: invalidArgErrors.slice(0, 3).map(r => r.errorMessage),
      });
    }

    // 模式 2: InvalidReference 错误
    const invalidRefErrors = reports.filter(r => r.errorType === 'InvalidReference');
    if (invalidRefErrors.length > 0) {
      patterns.push({
        pattern: 'InvalidReference',
        description: '无效引用错误',
        frequency: invalidRefErrors.length,
        affectedTestCases: Array.from(new Set(invalidRefErrors.map(r => r.testCaseId).filter(Boolean))) as string[],
        examples: invalidRefErrors.slice(0, 3).map(r => r.errorMessage),
      });
    }

    // 模式 3: API 未找到
    const apiNotFoundErrors = reports.filter(r => r.errorType === 'ApiNotFound');
    if (apiNotFoundErrors.length > 0) {
      patterns.push({
        pattern: 'ApiNotFound',
        description: 'API 不可用错误',
        frequency: apiNotFoundErrors.length,
        affectedTestCases: Array.from(new Set(apiNotFoundErrors.map(r => r.testCaseId).filter(Boolean))) as string[],
        examples: apiNotFoundErrors.slice(0, 3).map(r => r.errorMessage),
      });
    }

    // 模式 4: 特定功能的高频错误
    const errorsByTestCase = new Map<string, ErrorReport[]>();
    for (const report of reports) {
      if (report.testCaseId) {
        if (!errorsByTestCase.has(report.testCaseId)) {
          errorsByTestCase.set(report.testCaseId, []);
        }
        errorsByTestCase.get(report.testCaseId)!.push(report);
      }
    }

    const highFreqTestCases = Array.from(errorsByTestCase.entries())
      .filter(([_, errors]) => errors.length >= 3)
      .sort((a, b) => b[1].length - a[1].length)
      .slice(0, 5);

    for (const [testCaseId, errors] of highFreqTestCases) {
      patterns.push({
        pattern: `HighFrequency:${testCaseId}`,
        description: `测试用例 ${testCaseId} 频繁失败`,
        frequency: errors.length,
        affectedTestCases: [testCaseId],
        examples: errors.slice(0, 3).map(r => r.errorMessage),
      });
    }

    return patterns;
  }

  /**
   * 生成优化建议
   */
  private static generateRecommendations(reports: ErrorReport[]): Recommendation[] {
    const recommendations: Recommendation[] = [];
    const patterns = this.identifyPatterns(reports);

    // 基于错误模式生成建议
    for (const pattern of patterns) {
      if (pattern.pattern === 'InvalidArgument') {
        recommendations.push({
          priority: 'high',
          area: '参数验证',
          issue: `发现 ${pattern.frequency} 个 InvalidArgument 错误`,
          suggestion: '在 TOOLS.md 模板中添加参数验证代码，使用 getItemOrNullObject 检查对象是否存在',
          affectedCount: pattern.frequency,
        });
      }

      if (pattern.pattern === 'InvalidReference') {
        recommendations.push({
          priority: 'high',
          area: '引用检查',
          issue: `发现 ${pattern.frequency} 个 InvalidReference 错误`,
          suggestion: '在代码模板中添加引用存在性检查，避免访问不存在的对象',
          affectedCount: pattern.frequency,
        });
      }

      if (pattern.pattern === 'ApiNotFound') {
        recommendations.push({
          priority: 'medium',
          area: 'API 兼容性',
          issue: `发现 ${pattern.frequency} 个 ApiNotFound 错误`,
          suggestion: '在 SKILL.md 中标注 API 的平台支持情况，或提供替代方案',
          affectedCount: pattern.frequency,
        });
      }
    }

    // 基于统计信息生成建议
    const stats = this.generateStatistics(reports);
    for (const [category, count] of Object.entries(stats.errorsByCategory)) {
      if (count >= 5) {
        recommendations.push({
          priority: 'medium',
          area: category,
          issue: `${category} 类别中有 ${count} 个错误`,
          suggestion: `重点检查 ${category} 相关的代码模板，增强错误处理和参数验证`,
          affectedCount: count,
        });
      }
    }

    return recommendations.sort((a, b) => {
      const priorityOrder = { high: 0, medium: 1, low: 2 };
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    });
  }

  /**
   * 识别问题区域
   */
  private static identifyProblematicAreas(reports: ErrorReport[]): ProblematicArea[] {
    const areaMap = new Map<string, {
      errors: ErrorReport[];
      errorMessages: Map<string, number>;
    }>();

    for (const report of reports) {
      if (report.testCaseId) {
        const category = this.extractCategory(report.testCaseId);
        
        if (!areaMap.has(category)) {
          areaMap.set(category, {
            errors: [],
            errorMessages: new Map(),
          });
        }

        const area = areaMap.get(category)!;
        area.errors.push(report);
        area.errorMessages.set(
          report.errorMessage,
          (area.errorMessages.get(report.errorMessage) || 0) + 1
        );
      }
    }

    const areas: ProblematicArea[] = [];
    for (const [category, data] of areaMap.entries()) {
      const topErrors = Array.from(data.errorMessages.entries())
        .map(([message, count]) => ({ message, count }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

      areas.push({
        category,
        errorCount: data.errors.length,
        errorRate: data.errors.length / reports.length,
        topErrors,
      });
    }

    return areas.sort((a, b) => b.errorCount - a.errorCount);
  }

  /**
   * 从 testCaseId 提取类别
   */
  private static extractCategory(testCaseId: string): string {
    // testCaseId 格式: excel-001-创建工作表
    const parts = testCaseId.split('-');
    if (parts.length >= 3) {
      return parts.slice(2).join('-');
    }
    return 'unknown';
  }

  /**
   * 生成 Markdown 报告
   */
  static generateMarkdownReport(analysis: AnalysisResult): string {
    const lines: string[] = [];
    
    lines.push('# 错误分析报告\n');
    lines.push(`生成时间: ${new Date().toLocaleString()}\n`);
    
    // 摘要
    lines.push('## 摘要\n');
    lines.push(`- 总错误数: ${analysis.summary.totalErrors}`);
    lines.push(`- 受影响的测试用例: ${analysis.summary.uniqueTestCases}`);
    lines.push(`- 时间范围: ${analysis.summary.timeRange.start} ~ ${analysis.summary.timeRange.end}\n`);
    
    // 统计
    lines.push('## 错误统计\n');
    lines.push('### 按错误类型\n');
    for (const [type, count] of Object.entries(analysis.statistics.errorsByType)) {
      lines.push(`- ${type}: ${count} 次`);
    }
    lines.push('');
    
    lines.push('### 按功能类别\n');
    for (const [category, count] of Object.entries(analysis.statistics.errorsByCategory)) {
      lines.push(`- ${category}: ${count} 次`);
    }
    lines.push('');
    
    // 错误模式
    lines.push('## 错误模式\n');
    for (const pattern of analysis.patterns) {
      lines.push(`### ${pattern.pattern}\n`);
      lines.push(`${pattern.description}`);
      lines.push(`- 频率: ${pattern.frequency} 次`);
      lines.push(`- 影响的测试用例: ${pattern.affectedTestCases.length} 个`);
      lines.push('- 示例:');
      for (const example of pattern.examples.slice(0, 3)) {
        lines.push(`  - ${example}`);
      }
      lines.push('');
    }
    
    // 优化建议
    lines.push('## 优化建议\n');
    for (const rec of analysis.recommendations) {
      const priorityEmoji = rec.priority === 'high' ? '🔴' : rec.priority === 'medium' ? '🟡' : '🟢';
      lines.push(`### ${priorityEmoji} ${rec.area}\n`);
      lines.push(`**问题**: ${rec.issue}\n`);
      lines.push(`**建议**: ${rec.suggestion}\n`);
      lines.push(`**影响范围**: ${rec.affectedCount} 个错误\n`);
    }
    
    // 问题区域
    lines.push('## 问题区域排名\n');
    for (let i = 0; i < analysis.topProblematicAreas.length; i++) {
      const area = analysis.topProblematicAreas[i];
      lines.push(`### ${i + 1}. ${area.category}\n`);
      lines.push(`- 错误数: ${area.errorCount}`);
      lines.push(`- 错误率: ${(area.errorRate * 100).toFixed(1)}%`);
      lines.push('- 主要错误:');
      for (const error of area.topErrors) {
        lines.push(`  - ${error.message} (${error.count} 次)`);
      }
      lines.push('');
    }
    
    return lines.join('\n');
  }
}
