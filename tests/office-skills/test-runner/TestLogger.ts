/**
 * 测试日志记录器
 */

import type { TestResult, TestSession, ErrorReport } from './types';

export class TestLogger {
  private sessionId: string;
  private logs: string[] = [];

  constructor(sessionId: string) {
    this.sessionId = sessionId;
  }

  /**
   * 记录信息日志
   */
  info(message: string): void {
    const log = `[${new Date().toISOString()}] [INFO] ${message}`;
    this.logs.push(log);
    console.log(log);
  }

  /**
   * 记录错误日志
   */
  error(message: string, error?: any): void {
    const log = `[${new Date().toISOString()}] [ERROR] ${message}`;
    this.logs.push(log);
    console.error(log, error);
  }

  /**
   * 记录警告日志
   */
  warn(message: string): void {
    const log = `[${new Date().toISOString()}] [WARN] ${message}`;
    this.logs.push(log);
    console.warn(log);
  }

  /**
   * 记录测试开始
   */
  logTestStart(testCaseId: string, testName: string): void {
    this.info(`开始测试: ${testCaseId} - ${testName}`);
  }

  /**
   * 记录测试完成
   */
  logTestComplete(result: TestResult): void {
    const status = result.status === 'passed' ? '✅ 通过' : '❌ 失败';
    const duration = result.duration ? ` (${result.duration}ms)` : '';
    this.info(`测试完成: ${result.testCaseId} ${status}${duration}`);
    
    if (result.error) {
      this.error(`  错误: ${result.error.message}`);
    }
  }

  /**
   * 记录测试会话摘要
   */
  logSessionSummary(session: TestSession): void {
    this.info('='.repeat(60));
    this.info('测试会话摘要');
    this.info(`会话 ID: ${session.id}`);
    this.info(`应用类型: ${session.hostType.toUpperCase()}`);
    this.info(`开始时间: ${session.startTime}`);
    this.info(`结束时间: ${session.endTime || '进行中'}`);
    this.info('-'.repeat(60));
    this.info(`总测试数: ${session.summary.total}`);
    this.info(`✅ 通过: ${session.summary.passed}`);
    this.info(`❌ 失败: ${session.summary.failed}`);
    this.info(`⏭️  跳过: ${session.summary.skipped}`);
    this.info(`📊 错误率: ${(session.summary.errorRate * 100).toFixed(2)}%`);
    this.info('='.repeat(60));
  }

  /**
   * 保存错误报告到本地存储
   */
  async saveErrorReport(report: ErrorReport): Promise<void> {
    try {
      // 使用 localStorage 保存错误报告
      const key = `error-report-${report.timestamp}`;
      localStorage.setItem(key, JSON.stringify(report));
      
      // 维护错误报告索引
      const indexKey = 'error-reports-index';
      const indexData = localStorage.getItem(indexKey);
      const index = indexData ? JSON.parse(indexData) : [];
      index.push({
        key,
        timestamp: report.timestamp,
        testCaseId: report.testCaseId,
        errorType: report.errorType,
      });
      localStorage.setItem(indexKey, JSON.stringify(index));
      
      this.info(`错误报告已保存: ${key}`);
    } catch (error) {
      this.error('保存错误报告失败', error);
    }
  }

  /**
   * 导出所有日志
   */
  exportLogs(): string {
    return this.logs.join('\n');
  }

  /**
   * 清空日志
   */
  clearLogs(): void {
    this.logs = [];
  }

  /**
   * 获取所有错误报告
   */
  static getErrorReports(): ErrorReport[] {
    try {
      const indexKey = 'error-reports-index';
      const indexData = localStorage.getItem(indexKey);
      if (!indexData) return [];
      
      const index = JSON.parse(indexData);
      const reports: ErrorReport[] = [];
      
      for (const item of index) {
        const reportData = localStorage.getItem(item.key);
        if (reportData) {
          reports.push(JSON.parse(reportData));
        }
      }
      
      return reports;
    } catch (error) {
      console.error('获取错误报告失败', error);
      return [];
    }
  }

  /**
   * 清除所有错误报告
   */
  static clearErrorReports(): void {
    try {
      const indexKey = 'error-reports-index';
      const indexData = localStorage.getItem(indexKey);
      if (!indexData) return;
      
      const index = JSON.parse(indexData);
      for (const item of index) {
        localStorage.removeItem(item.key);
      }
      localStorage.removeItem(indexKey);
      
      console.log('所有错误报告已清除');
    } catch (error) {
      console.error('清除错误报告失败', error);
    }
  }
}
