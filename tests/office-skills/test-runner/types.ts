/**
 * 测试系统类型定义
 */

export type HostType = 'excel' | 'word' | 'powerpoint';

export type TestStatus = 'pending' | 'running' | 'passed' | 'failed' | 'skipped';

export type ErrorType = 
  | 'InvalidArgument' 
  | 'InvalidReference' 
  | 'ApiNotFound' 
  | 'GeneralException'
  | 'NetworkError'
  | 'TimeoutError'
  | 'UnknownError';

/**
 * 测试用例定义
 */
export interface TestCase {
  id: string;
  category: string;
  name: string;
  description: string;
  userInput: string;
  expectedCode?: string;
  expectedBehavior: string;
  validationSteps: string[];
  toolsTemplate: string;
  priority?: 'high' | 'medium' | 'low';
  tags?: string[];
}

/**
 * 测试结果
 */
export interface TestResult {
  testCaseId: string;
  status: TestStatus;
  startTime: string;
  endTime?: string;
  duration?: number;
  error?: TestError;
  actualCode?: string;
  logs?: string[];
}

/**
 * 测试错误
 */
export interface TestError {
  type: ErrorType;
  code?: string;
  message: string;
  stackTrace?: string;
  debugInfo?: any;
}

/**
 * 错误报告
 */
export interface ErrorReport {
  timestamp: string;
  testCaseId: string;
  hostType: HostType;
  errorType: ErrorType;
  errorCode?: string;
  errorMessage: string;
  stackTrace?: string;
  userInput: string;
  generatedCode: string;
  context: {
    officeVersion: string;
    platform: string;
    browserInfo?: string;
  };
}

/**
 * 测试套件
 */
export interface TestSuite {
  id: string;
  name: string;
  hostType: HostType;
  testCases: TestCase[];
  metadata?: {
    version: string;
    generatedAt: string;
    sourceFile: string;
  };
}

/**
 * 测试会话
 */
export interface TestSession {
  id: string;
  hostType: HostType;
  startTime: string;
  endTime?: string;
  results: TestResult[];
  summary: {
    total: number;
    passed: number;
    failed: number;
    skipped: number;
    errorRate: number;
  };
}

/**
 * 错误统计
 */
export interface ErrorStatistics {
  totalErrors: number;
  errorsByType: Record<ErrorType, number>;
  errorsByCategory: Record<string, number>;
  topErrors: Array<{
    message: string;
    count: number;
    testCases: string[];
  }>;
}
