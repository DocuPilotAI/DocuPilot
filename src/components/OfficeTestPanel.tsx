"use client";

import { useEffect, useMemo, useState } from "react";
import type { HostType, TestCase, TestSession } from "../../tests/office-skills/test-runner/types";
import { TestExecutor } from "../../tests/office-skills/test-runner/TestExecutor";
import { TestLogger } from "../../tests/office-skills/test-runner/TestLogger";

type Props = {
  hostType: HostType;
};

export function OfficeTestPanel({ hostType }: Props) {
  const [testCases, setTestCases] = useState<TestCase[]>([]);
  const [loading, setLoading] = useState(false);
  const [session, setSession] = useState<TestSession | null>(null);
  const [progress, setProgress] = useState({ current: 0, total: 0 });
  const [loadError, setLoadError] = useState<string | null>(null);
  const [useDebugCases, setUseDebugCases] = useState(false);

  const officeStatus = useMemo(() => {
    const g = globalThis as any;
    const officeLoaded = typeof g.Office !== "undefined";
    const apiLoaded =
      hostType === "excel"
        ? typeof g.Excel !== "undefined"
        : hostType === "word"
          ? typeof g.Word !== "undefined"
          : typeof g.PowerPoint !== "undefined";
    const version = officeLoaded ? g.Office?.context?.diagnostics?.version || "unknown" : "";
    return { officeLoaded, apiLoaded, version };
  }, [hostType]);

  useEffect(() => {
    let cancelled = false;

    async function loadSuite() {
      setLoadError(null);
      setTestCases([]);
      try {
        const suffix = useDebugCases ? '-debug' : '';
        const res = await fetch(`/test-cases/${hostType}-test-cases${suffix}.json`, { cache: "no-store" });
        if (!res.ok) {
          throw new Error(`加载测试用例失败: HTTP ${res.status}`);
        }
        const data = await res.json();
        const cases: TestCase[] = Array.isArray(data) ? data : (data.testCases ?? []);
        if (!cancelled) setTestCases(cases);
      } catch (e) {
        if (!cancelled) {
          setLoadError(e instanceof Error ? e.message : String(e));
        }
      }
    }

    loadSuite();
    return () => {
      cancelled = true;
    };
  }, [hostType, useDebugCases]);

  const runRegression = async () => {
    if (!officeStatus.officeLoaded || !officeStatus.apiLoaded) {
      alert("当前未检测到完整 Office 环境（Office.js 或对应宿主 API 未加载）。请确认在任务窗格中运行。");
      return;
    }
    if (testCases.length === 0) {
      alert("没有可运行的测试用例。");
      return;
    }

    setLoading(true);
    setSession(null);
    setProgress({ current: 0, total: 0 });

    try {
      const executor = new TestExecutor(hostType);
      const result = await executor.executeTestSuite(testCases, (current, total) =>
        setProgress({ current, total })
      );
      setSession(result);
    } catch (e) {
      console.error("[OfficeTestPanel] 执行测试失败:", e);
      alert("执行测试失败: " + (e instanceof Error ? e.message : String(e)));
    } finally {
      setLoading(false);
      setProgress({ current: 0, total: 0 });
    }
  };

  const exportResults = () => {
    if (!session) return;
    const data = JSON.stringify(session, null, 2);
    const blob = new Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `taskpane-test-results-${hostType}-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportErrors = () => {
    const errors = TestLogger.getErrorReports();
    const data = JSON.stringify(errors, null, 2);
    const blob = new Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `taskpane-error-reports-${hostType}-${Date.now()}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const clearErrors = () => {
    if (confirm("确定要清除所有错误报告吗?")) {
      TestLogger.clearErrorReports();
      alert("错误报告已清除");
    }
  };

  return (
    <div className="h-full w-full overflow-y-auto bg-gray-50 p-4">
      {/* 环境状态 */}
      {!officeStatus.officeLoaded && (
        <div className="bg-yellow-50 border-l-4 border-yellow-400 p-3 mb-4">
          <div className="text-sm text-yellow-800 font-medium">⚠️ Office.js 未加载</div>
          <div className="text-xs text-yellow-700 mt-1">
            请确认你是在 Office 任务窗格中打开此页面（而不是普通浏览器）。
          </div>
        </div>
      )}

      {officeStatus.officeLoaded && !officeStatus.apiLoaded && (
        <div className="bg-blue-50 border-l-4 border-blue-400 p-3 mb-4">
          <div className="text-sm text-blue-800 font-medium">ℹ️ Office.js 已加载，但宿主 API 未加载</div>
          <div className="text-xs text-blue-700 mt-1">
            Office 版本: {officeStatus.version} | 当前宿主: {hostType.toUpperCase()}
          </div>
        </div>
      )}

      {officeStatus.officeLoaded && officeStatus.apiLoaded && (
        <div className="bg-green-50 border-l-4 border-green-400 p-3 mb-4">
          <div className="text-sm text-green-800 font-medium">✅ Office 环境正常</div>
          <div className="text-xs text-green-700 mt-1">
            Office 版本: {officeStatus.version} | 宿主: {hostType.toUpperCase()}
          </div>
        </div>
      )}

      {/* 控制区 */}
      <div className="bg-white rounded-lg shadow p-4 mb-4">
        <div className="flex items-center justify-between gap-2 mb-3">
          <div>
            <div className="text-sm font-medium">回归测试</div>
            <div className="text-xs text-gray-600 mt-1">用例数: {testCases.length}</div>
            {loadError && <div className="text-xs text-red-600 mt-1">{loadError}</div>}
          </div>

          <div className="flex gap-2 items-center">
            {/* 调试模式切换 */}
            <label className="flex items-center gap-2 text-xs text-gray-700 cursor-pointer">
              <input
                type="checkbox"
                checked={useDebugCases}
                onChange={(e) => setUseDebugCases(e.target.checked)}
                disabled={loading}
                className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
              />
              <span className="select-none">仅测试问题用例 (4个)</span>
            </label>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <div className="flex items-center gap-2">
            <button
              onClick={runRegression}
              disabled={loading || testCases.length === 0}
              className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-sm"
            >
              {loading ? "测试中..." : "运行回归"}
            </button>
            <button
              onClick={exportResults}
              disabled={!session}
              className="px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 disabled:bg-gray-400 disabled:cursor-not-allowed text-sm"
            >
              导出结果
            </button>
            <button
              onClick={exportErrors}
              className="px-4 py-2 bg-orange-600 text-white rounded-md hover:bg-orange-700 text-sm"
            >
              导出错误
            </button>
            <button
              onClick={clearErrors}
              className="px-4 py-2 bg-red-600 text-white rounded-md hover:bg-red-700 text-sm"
            >
              清空错误
            </button>
          </div>
        </div>

        {loading && progress.total > 0 && (
          <div className="mt-3">
            <div className="flex justify-between text-xs mb-1">
              <span>进度</span>
              <span>
                {progress.current} / {progress.total}
              </span>
            </div>
            <progress
              className="w-full h-2 rounded-full overflow-hidden"
              value={progress.current}
              max={progress.total}
            />
          </div>
        )}
      </div>

      {/* 结果区 */}
      {session && (
        <div className="bg-white rounded-lg shadow p-4">
          <div className="text-sm font-medium mb-3">测试结果</div>

          <div className="grid grid-cols-2 md:grid-cols-5 gap-3 mb-4">
            <div className="bg-gray-50 p-3 rounded">
              <div className="text-xs text-gray-600">总数</div>
              <div className="text-xl font-bold">{session.summary.total}</div>
            </div>
            <div className="bg-green-50 p-3 rounded">
              <div className="text-xs text-gray-600">✅ 通过</div>
              <div className="text-xl font-bold text-green-600">{session.summary.passed}</div>
            </div>
            <div className="bg-red-50 p-3 rounded">
              <div className="text-xs text-gray-600">❌ 失败</div>
              <div className="text-xl font-bold text-red-600">{session.summary.failed}</div>
            </div>
            <div className="bg-gray-50 p-3 rounded">
              <div className="text-xs text-gray-600">⏭️ 跳过</div>
              <div className="text-xl font-bold">{session.summary.skipped}</div>
            </div>
            <div className="bg-blue-50 p-3 rounded">
              <div className="text-xs text-gray-600">错误率</div>
              <div className="text-xl font-bold text-blue-600">
                {(session.summary.errorRate * 100).toFixed(1)}%
              </div>
            </div>
          </div>

          <div className="space-y-2 max-h-96 overflow-y-auto">
            {session.results.map((result) => {
              const tc = testCases.find((t) => t.id === result.testCaseId);
              const ok = result.status === "passed";
              return (
                <div
                  key={result.testCaseId}
                  className={`border-l-4 p-3 rounded ${
                    ok ? "border-green-500 bg-green-50" : "border-red-500 bg-red-50"
                  }`}
                >
                  <div className="flex justify-between items-start">
                    <div>
                      <div className="text-sm font-medium">
                        {ok ? "✅" : "❌"} {tc?.name || result.testCaseId}
                      </div>
                      {tc?.category && <div className="text-xs text-gray-600">{tc.category}</div>}
                    </div>
                    <div className="text-xs text-gray-500">{result.duration ?? 0}ms</div>
                  </div>
                  {result.error && (
                    <div className="mt-2 text-xs text-red-700 whitespace-pre-wrap">
                      {result.error.type}: {result.error.message}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

