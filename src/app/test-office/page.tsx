"use client";

import { useState, useEffect } from 'react';
import type { TestCase, TestSession, HostType } from '../../../tests/office-skills/test-runner/types';
import { TestExecutor } from '../../../tests/office-skills/test-runner/TestExecutor';
import { TestLogger } from '../../../tests/office-skills/test-runner/TestLogger';
import { isInOfficeEnvironment, loadOfficeJs, waitForOfficeReady } from '@/lib/office-loader';

export default function TestOfficePage() {
  const [hostType, setHostType] = useState<HostType>('excel');
  const [testCases, setTestCases] = useState<TestCase[]>([]);
  const [loading, setLoading] = useState(false);
  const [session, setSession] = useState<TestSession | null>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [officeEnvironment, setOfficeEnvironment] = useState({
    officeLoaded: false,
    apiLoaded: false,
    version: '',
  });

  // 支持通过 URL 强制指定 host（用于“按钮硬编码入口”）
  useEffect(() => {
    try {
      const params = new URLSearchParams(window.location.search);
      const host = params.get('host');
      if (host === 'excel' || host === 'word' || host === 'powerpoint') {
        setHostType(host);
      }
    } catch {
      // ignore
    }
  }, []);

  const ensureOfficeReady = async (): Promise<boolean> => {
    try {
      await loadOfficeJs();
      const host = await waitForOfficeReady({ timeoutMs: 10000 });
      // host 为 null 时，通常表示不在 Add-in 里或 Office.onReady 没触发
      return Boolean(host) || isInOfficeEnvironment();
    } catch {
      return false;
    }
  };

  // 检查 Office 环境
  useEffect(() => {
    let cancelled = false;

    const checkEnvironment = () => {
      const g = globalThis as any;
      const officeLoaded = typeof g.Office !== 'undefined';
      let apiLoaded = false;
      let version = '';

      if (officeLoaded) {
        version = g.Office.context?.diagnostics?.version || 'unknown';
        if (hostType === 'excel') {
          apiLoaded = typeof g.Excel !== 'undefined';
        } else if (hostType === 'word') {
          apiLoaded = typeof g.Word !== 'undefined';
        } else if (hostType === 'powerpoint') {
          apiLoaded = typeof g.PowerPoint !== 'undefined';
        }
      }

      setOfficeEnvironment({
        officeLoaded,
        apiLoaded,
        version,
      });
    };

    const initOffice = async () => {
      await ensureOfficeReady();
      if (!cancelled) {
        checkEnvironment();
      }
    };

    initOffice();

    return () => {
      cancelled = true;
    };
  }, [hostType]);

  // 加载测试用例
  useEffect(() => {
    loadTestCases(hostType);
  }, [hostType]);

  const loadTestCases = async (host: HostType) => {
    try {
      const response = await fetch(`/test-cases/${host}-test-cases.json`);
      if (response.ok) {
        const data = await response.json();
        setTestCases(data.testCases || data);
      } else {
        console.error('加载测试用例失败: HTTP', response.status);
        setTestCases([]);
      }
    } catch (error) {
      console.error('加载测试用例失败:', error);
      setTestCases([]);
    }
  };

  const categories = ['all', ...Array.from(new Set(testCases.map(tc => tc.category)))];

  const filteredTestCases = selectedCategory === 'all' 
    ? testCases 
    : testCases.filter(tc => tc.category === selectedCategory);

  const runTests = async () => {
    const ready = await ensureOfficeReady();
    if (!ready) {
      alert('Office.js 未加载或尚未就绪，请在 Office Add-in 任务窗格中打开此页面，并稍后再试。');
      return;
    }

    setLoading(true);
    setSession(null);
    
    try {
      const executor = new TestExecutor(hostType);
      const result = await executor.executeTestSuite(
        filteredTestCases,
        (current, total) => setProgress({ current, total })
      );
      setSession(result);
    } catch (error) {
      console.error('执行测试失败:', error);
      alert('执行测试失败: ' + (error instanceof Error ? error.message : String(error)));
    } finally {
      setLoading(false);
      setProgress({ current: 0, total: 0 });
    }
  };

  const exportResults = () => {
    if (!session) return;
    
    const data = JSON.stringify(session, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `test-results-${hostType}-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportErrors = () => {
    const errors = TestLogger.getErrorReports();
    const data = JSON.stringify(errors, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `error-reports-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const clearErrors = () => {
    if (confirm('确定要清除所有错误报告吗?')) {
      TestLogger.clearErrorReports();
      alert('错误报告已清除');
    }
  };

  return (
    <div className="h-screen overflow-y-auto overflow-x-hidden bg-gray-50 p-8">
      <div className="max-w-7xl mx-auto pb-8">
        <h1 className="text-3xl font-bold mb-8">Office Skills 测试系统</h1>

        {/* 环境检查提示 */}
        {!officeEnvironment.officeLoaded && (
          <div className="bg-yellow-50 border-l-4 border-yellow-400 p-4 mb-6">
            <div className="flex">
              <div className="flex-shrink-0">
                <svg className="h-5 w-5 text-yellow-400" viewBox="0 0 20 20" fill="currentColor">
                  <path fillRule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
                </svg>
              </div>
              <div className="ml-3">
                <h3 className="text-sm font-medium text-yellow-800">
                  ⚠️ Office 环境未检测到
                </h3>
                <div className="mt-2 text-sm text-yellow-700">
                  <p>此测试系统必须在 Office Add-in 环境中运行。请：</p>
                  <ol className="list-decimal list-inside mt-2 space-y-1">
                    <li>在 Excel/Word/PowerPoint 中加载 Add-in (manifest.xml)</li>
                    <li>在 Add-in 的任务窗格中访问此页面</li>
                    <li>确保 Office.js 已正确加载</li>
                  </ol>
                  <p className="mt-2 text-xs">
                    详细说明请查看: <code className="bg-yellow-100 px-1">tests/office-skills/TEST_PRINCIPLE.md</code>
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {officeEnvironment.officeLoaded && !officeEnvironment.apiLoaded && (
          <div className="bg-blue-50 border-l-4 border-blue-400 p-4 mb-6">
            <div className="flex">
              <div className="ml-3">
                <h3 className="text-sm font-medium text-blue-800">
                  ℹ️ Office.js 已加载，但应用 API 未加载
                </h3>
                <div className="mt-2 text-sm text-blue-700">
                  <p>Office 版本: {officeEnvironment.version}</p>
                  <p className="mt-1">
                    请确保在正确的应用中运行（当前选择: {hostType.toUpperCase()}）
                  </p>
                </div>
              </div>
            </div>
          </div>
        )}

        {officeEnvironment.officeLoaded && officeEnvironment.apiLoaded && (
          <div className="bg-green-50 border-l-4 border-green-400 p-4 mb-6">
            <div className="flex">
              <div className="ml-3">
                <h3 className="text-sm font-medium text-green-800">
                  ✅ Office 环境正常
                </h3>
                <div className="mt-1 text-sm text-green-700">
                  Office 版本: {officeEnvironment.version} | 应用: {hostType.toUpperCase()} API 已加载
                </div>
              </div>
            </div>
          </div>
        )}

        {/* 控制面板 */}
        <div className="bg-white rounded-lg shadow p-6 mb-6">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div>
              <label className="block text-sm font-medium mb-2">选择应用</label>
              <select
                value={hostType}
                onChange={(e) => setHostType(e.target.value as HostType)}
                disabled={loading}
                aria-label="选择应用"
                title="选择应用"
                className="w-full px-3 py-2 border rounded-md"
              >
                <option value="excel">Excel</option>
                <option value="word">Word</option>
                <option value="powerpoint">PowerPoint</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium mb-2">选择类别</label>
              <select
                value={selectedCategory}
                onChange={(e) => setSelectedCategory(e.target.value)}
                disabled={loading}
                aria-label="选择类别"
                title="选择类别"
                className="w-full px-3 py-2 border rounded-md"
              >
                {categories.map(cat => (
                  <option key={cat} value={cat}>
                    {cat === 'all' ? '全部' : cat}
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium mb-2">测试用例数</label>
              <div className="text-2xl font-bold text-blue-600">
                {filteredTestCases.length}
              </div>
            </div>
          </div>

          <div className="flex gap-2">
            <button
              onClick={runTests}
              disabled={loading || filteredTestCases.length === 0}
              className="px-6 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {loading ? '测试中...' : '开始测试'}
            </button>
            
            <button
              onClick={exportResults}
              disabled={!session}
              className="px-6 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              导出结果
            </button>

            <button
              onClick={exportErrors}
              className="px-6 py-2 bg-orange-600 text-white rounded-md hover:bg-orange-700"
            >
              导出错误
            </button>

            <button
              onClick={clearErrors}
              className="px-6 py-2 bg-red-600 text-white rounded-md hover:bg-red-700"
            >
              清除错误
            </button>
          </div>

          {/* 进度条 */}
          {loading && progress.total > 0 && (
            <div className="mt-4">
              <div className="flex justify-between text-sm mb-1">
                <span>测试进度</span>
                <span>{progress.current} / {progress.total}</span>
              </div>
              <div className="w-full bg-gray-200 rounded-full h-2">
                <div
                  className="bg-blue-600 h-2 rounded-full transition-all"
                  style={{ width: `${(progress.current / progress.total) * 100}%` }}
                />
              </div>
            </div>
          )}
        </div>

        {/* 测试结果 */}
        {session && (
          <div className="bg-white rounded-lg shadow p-6 mb-6">
            <h2 className="text-xl font-bold mb-4">测试结果</h2>
            
            <div className="grid grid-cols-2 md:grid-cols-5 gap-4 mb-4">
              <div className="bg-gray-50 p-4 rounded">
                <div className="text-sm text-gray-600">总测试数</div>
                <div className="text-2xl font-bold">{session.summary.total}</div>
              </div>
              <div className="bg-green-50 p-4 rounded">
                <div className="text-sm text-gray-600">✅ 通过</div>
                <div className="text-2xl font-bold text-green-600">{session.summary.passed}</div>
              </div>
              <div className="bg-red-50 p-4 rounded">
                <div className="text-sm text-gray-600">❌ 失败</div>
                <div className="text-2xl font-bold text-red-600">{session.summary.failed}</div>
              </div>
              <div className="bg-gray-50 p-4 rounded">
                <div className="text-sm text-gray-600">⏭️ 跳过</div>
                <div className="text-2xl font-bold">{session.summary.skipped}</div>
              </div>
              <div className="bg-blue-50 p-4 rounded">
                <div className="text-sm text-gray-600">错误率</div>
                <div className="text-2xl font-bold text-blue-600">
                  {(session.summary.errorRate * 100).toFixed(1)}%
                </div>
              </div>
            </div>

            {/* 详细结果列表 */}
            <div className="space-y-2 max-h-96 overflow-y-auto">
              {session.results.map((result) => {
                const testCase = testCases.find(tc => tc.id === result.testCaseId);
                return (
                  <div
                    key={result.testCaseId}
                    className={`border-l-4 p-3 rounded ${
                      result.status === 'passed'
                        ? 'border-green-500 bg-green-50'
                        : 'border-red-500 bg-red-50'
                    }`}
                  >
                    <div className="flex justify-between items-start">
                      <div>
                        <div className="font-medium">
                          {result.status === 'passed' ? '✅' : '❌'} {testCase?.name || result.testCaseId}
                        </div>
                        {testCase && (
                          <div className="text-sm text-gray-600">{testCase.category}</div>
                        )}
                      </div>
                      <div className="text-sm text-gray-500">
                        {result.duration}ms
                      </div>
                    </div>
                    {result.error && (
                      <div className="mt-2 text-sm text-red-600">
                        错误: {result.error.message}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* 测试用例列表 */}
        <div className="bg-white rounded-lg shadow p-6">
          <h2 className="text-xl font-bold mb-4">测试用例列表</h2>
          <div className="space-y-2 max-h-96 overflow-y-auto">
            {filteredTestCases.length === 0 ? (
              <div className="text-center py-8 text-gray-500">
                暂无测试用例。请先运行 <code className="bg-gray-100 px-2 py-1 rounded">generate-test-cases.ts</code> 脚本生成测试用例。
              </div>
            ) : (
              filteredTestCases.map((testCase) => (
                <div key={testCase.id} className="border rounded p-3 hover:bg-gray-50">
                  <div className="font-medium">{testCase.name}</div>
                  <div className="text-sm text-gray-600">{testCase.description}</div>
                  <div className="flex gap-2 mt-2">
                    <span className="text-xs bg-blue-100 text-blue-800 px-2 py-1 rounded">
                      {testCase.category}
                    </span>
                    {testCase.priority && (
                      <span className={`text-xs px-2 py-1 rounded ${
                        testCase.priority === 'high' ? 'bg-red-100 text-red-800' :
                        testCase.priority === 'medium' ? 'bg-yellow-100 text-yellow-800' :
                        'bg-gray-100 text-gray-800'
                      }`}>
                        {testCase.priority}
                      </span>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
