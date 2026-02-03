#!/usr/bin/env node
/**
 * 分析错误并生成优化建议
 */

import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 注意：这个脚本需要在浏览器环境中运行才能访问 localStorage
// 这里提供一个命令行版本，从文件读取错误报告

interface ErrorReport {
  timestamp: string;
  testCaseId?: string;
  hostType: 'excel' | 'word' | 'powerpoint';
  errorType: string;
  errorCode?: string;
  errorMessage: string;
  stackTrace?: string;
  userInput?: string;
  generatedCode: string;
  context: {
    officeVersion: string;
    platform: string;
    browserInfo?: string;
  };
}

/**
 * 从 JSON 文件加载错误报告
 */
function loadErrorReports(filePath: string): ErrorReport[] {
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    return JSON.parse(content);
  } catch (error) {
    console.error('加载错误报告失败:', error);
    return [];
  }
}

/**
 * 分析错误报告
 */
function analyzeReports(reports: ErrorReport[]) {
  if (reports.length === 0) {
    console.log('没有错误报告需要分析');
    return;
  }

  console.log('='.repeat(60));
  console.log('错误分析报告');
  console.log('='.repeat(60));
  console.log();

  // 统计
  console.log('📊 统计信息');
  console.log('-'.repeat(60));
  console.log(`总错误数: ${reports.length}`);
  
  const uniqueTestCases = new Set(reports.map(r => r.testCaseId).filter(Boolean)).size;
  console.log(`受影响的测试用例: ${uniqueTestCases}`);
  
  // 按错误类型统计
  const errorsByType: Record<string, number> = {};
  for (const report of reports) {
    errorsByType[report.errorType] = (errorsByType[report.errorType] || 0) + 1;
  }
  
  console.log('\n按错误类型:');
  for (const [type, count] of Object.entries(errorsByType).sort((a, b) => b[1] - a[1])) {
    console.log(`  - ${type}: ${count} 次 (${((count / reports.length) * 100).toFixed(1)}%)`);
  }
  
  // 按主机类型统计
  const errorsByHost: Record<string, number> = {};
  for (const report of reports) {
    errorsByHost[report.hostType] = (errorsByHost[report.hostType] || 0) + 1;
  }
  
  console.log('\n按应用类型:');
  for (const [host, count] of Object.entries(errorsByHost).sort((a, b) => b[1] - a[1])) {
    console.log(`  - ${host.toUpperCase()}: ${count} 次`);
  }
  
  // Top 错误消息
  const errorMessages = new Map<string, number>();
  for (const report of reports) {
    errorMessages.set(report.errorMessage, (errorMessages.get(report.errorMessage) || 0) + 1);
  }
  
  const topErrors = Array.from(errorMessages.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  console.log('\n🔴 Top 10 错误消息:');
  for (let i = 0; i < topErrors.length; i++) {
    const [message, count] = topErrors[i];
    console.log(`\n${i + 1}. ${message}`);
    console.log(`   出现次数: ${count}`);
    
    // 找到受影响的测试用例
    const affectedTests = reports
      .filter(r => r.errorMessage === message && r.testCaseId)
      .map(r => r.testCaseId)
      .filter((v, i, a) => a.indexOf(v) === i)
      .slice(0, 5);
    
    if (affectedTests.length > 0) {
      console.log(`   受影响的测试: ${affectedTests.join(', ')}`);
    }
  }
  
  // 优化建议
  console.log('\n' + '='.repeat(60));
  console.log('💡 优化建议');
  console.log('='.repeat(60));
  
  let priority = 1;
  
  // InvalidArgument 错误
  const invalidArgCount = errorsByType['InvalidArgument'] || 0;
  if (invalidArgCount > 0) {
    console.log(`\n${priority++}. 参数验证问题 (${invalidArgCount} 个错误)`);
    console.log('   建议: 在 TOOLS.md 模板中添加参数验证');
    console.log('   - 使用 getItemOrNullObject 检查对象是否存在');
    console.log('   - 添加参数类型和范围检查');
    console.log('   - 在 SKILL.md 中补充参数说明');
  }
  
  // InvalidReference 错误
  const invalidRefCount = errorsByType['InvalidReference'] || 0;
  if (invalidRefCount > 0) {
    console.log(`\n${priority++}. 引用检查问题 (${invalidRefCount} 个错误)`);
    console.log('   建议: 增强引用存在性检查');
    console.log('   - 在访问对象前先验证是否存在');
    console.log('   - 提供更清晰的错误提示');
  }
  
  // ApiNotFound 错误
  const apiNotFoundCount = errorsByType['ApiNotFound'] || 0;
  if (apiNotFoundCount > 0) {
    console.log(`\n${priority++}. API 兼容性问题 (${apiNotFoundCount} 个错误)`);
    console.log('   建议: 标注 API 平台支持情况');
    console.log('   - 在 SKILL.md 中说明 API 的版本要求');
    console.log('   - 提供替代方案或降级处理');
  }
  
  console.log('\n' + '='.repeat(60));
  console.log('分析完成!');
  console.log('='.repeat(60));
}

/**
 * 主函数
 */
function main() {
  const args = process.argv.slice(2);
  
  if (args.length === 0) {
    console.log('用法: npx tsx analyze-errors.ts <error-reports.json>');
    console.log('\n示例:');
    console.log('  npx tsx analyze-errors.ts error-reports.json');
    console.log('  npx tsx analyze-errors.ts ../error-analysis/error-reports/latest.json');
    return;
  }
  
  const filePath = path.resolve(args[0]);
  
  if (!fs.existsSync(filePath)) {
    console.error(`错误: 文件不存在: ${filePath}`);
    return;
  }
  
  console.log(`\n读取错误报告: ${filePath}\n`);
  
  const reports = loadErrorReports(filePath);
  analyzeReports(reports);
}

// 运行
main();
