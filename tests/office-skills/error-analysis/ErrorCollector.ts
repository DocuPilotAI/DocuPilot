/**
 * 错误收集器 - 提供统一的错误收集接口
 */

import type { ErrorReport } from '../test-runner/types';

export class ErrorCollector {
  private static readonly INDEX_KEY = 'error-reports-index';

  /**
   * 保存错误报告
   */
  static saveErrorReport(report: ErrorReport): void {
    try {
      const key = `error-report-${report.timestamp}`;
      localStorage.setItem(key, JSON.stringify(report));
      
      // 更新索引
      const index = this.getIndex();
      index.push({
        key,
        timestamp: report.timestamp,
        testCaseId: report.testCaseId,
        errorType: report.errorType,
        hostType: report.hostType,
      });
      
      // 只保留最近 1000 条
      if (index.length > 1000) {
        const oldestKey = index.shift().key;
        localStorage.removeItem(oldestKey);
      }
      
      localStorage.setItem(this.INDEX_KEY, JSON.stringify(index));
    } catch (error) {
      console.error('保存错误报告失败:', error);
    }
  }

  /**
   * 获取所有错误报告
   */
  static getAllReports(): ErrorReport[] {
    try {
      const index = this.getIndex();
      const reports: ErrorReport[] = [];
      
      for (const item of index) {
        const reportData = localStorage.getItem(item.key);
        if (reportData) {
          reports.push(JSON.parse(reportData));
        }
      }
      
      return reports.sort((a, b) => 
        new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime()
      );
    } catch (error) {
      console.error('获取错误报告失败:', error);
      return [];
    }
  }

  /**
   * 按主机类型获取错误报告
   */
  static getReportsByHost(hostType: 'excel' | 'word' | 'powerpoint'): ErrorReport[] {
    return this.getAllReports().filter(r => r.hostType === hostType);
  }

  /**
   * 按错误类型获取错误报告
   */
  static getReportsByErrorType(errorType: string): ErrorReport[] {
    return this.getAllReports().filter(r => r.errorType === errorType);
  }

  /**
   * 按时间范围获取错误报告
   */
  static getReportsByTimeRange(startTime: Date, endTime: Date): ErrorReport[] {
    return this.getAllReports().filter(r => {
      const timestamp = new Date(r.timestamp);
      return timestamp >= startTime && timestamp <= endTime;
    });
  }

  /**
   * 清除所有错误报告
   */
  static clearAllReports(): void {
    try {
      const index = this.getIndex();
      for (const item of index) {
        localStorage.removeItem(item.key);
      }
      localStorage.removeItem(this.INDEX_KEY);
      console.log('所有错误报告已清除');
    } catch (error) {
      console.error('清除错误报告失败:', error);
    }
  }

  /**
   * 导出错误报告为 JSON
   */
  static exportReportsAsJSON(): string {
    const reports = this.getAllReports();
    return JSON.stringify(reports, null, 2);
  }

  /**
   * 获取错误统计
   */
  static getStatistics() {
    const reports = this.getAllReports();
    
    const stats = {
      total: reports.length,
      byErrorType: {} as Record<string, number>,
      byHostType: {} as Record<string, number>,
      recent24h: 0,
      recent7d: 0,
    };
    
    const now = new Date();
    const day24h = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const day7d = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    for (const report of reports) {
      // 按错误类型统计
      stats.byErrorType[report.errorType] = (stats.byErrorType[report.errorType] || 0) + 1;
      
      // 按主机类型统计
      stats.byHostType[report.hostType] = (stats.byHostType[report.hostType] || 0) + 1;
      
      // 时间范围统计
      const timestamp = new Date(report.timestamp);
      if (timestamp >= day24h) {
        stats.recent24h++;
      }
      if (timestamp >= day7d) {
        stats.recent7d++;
      }
    }
    
    return stats;
  }

  /**
   * 获取索引
   */
  private static getIndex(): any[] {
    try {
      const indexData = localStorage.getItem(this.INDEX_KEY);
      return indexData ? JSON.parse(indexData) : [];
    } catch {
      return [];
    }
  }
}
