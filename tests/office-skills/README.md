# Office Skills 测试系统

## 概述

这是一个半自动化测试系统,用于测试 Excel、Word、PowerPoint 的 Office.js 功能模板,收集错误并优化 Skills 提示语。

## 目录结构

```
tests/office-skills/
├── test-cases/              # 测试用例定义(JSON)
│   ├── excel-test-cases.json
│   ├── word-test-cases.json
│   └── powerpoint-test-cases.json
├── test-runner/             # 测试执行相关代码
│   ├── TestRunner.tsx       # React 测试界面组件
│   ├── TestExecutor.ts      # 测试执行逻辑
│   └── TestLogger.ts        # 测试日志记录
├── error-analysis/          # 错误分析系统
│   ├── ErrorCollector.ts    # 错误收集器
│   ├── ErrorAnalyzer.ts     # 错误分析器
│   └── error-reports/       # 错误报告存储目录
├── scripts/                 # 工具脚本
│   ├── generate-test-cases.ts  # 从 TOOLS.md 生成测试用例
│   └── analyze-errors.ts       # 分析错误并生成报告
└── README.md               # 本文件
```

## 使用方法

### 1. 生成测试用例

```bash
cd DocuPilot
npx tsx tests/office-skills/scripts/generate-test-cases.ts
```

这将从 `.claude/skills/*/TOOLS.md` 文件中自动生成测试用例。

### 2. 运行测试

1. 启动开发服务器: `npm run dev:https`
2. 在 Office 应用中加载 Add-in
3. 访问测试界面: `https://localhost:3000/test-office`
4. 选择应用类型(Excel/Word/PowerPoint)
5. 执行测试并查看结果

### 3. 分析错误

```bash
npx tsx tests/office-skills/scripts/analyze-errors.ts
```

这将分析收集到的错误数据,生成优化建议报告。

## 测试用例格式

每个测试用例包含以下字段:

```json
{
  "id": "excel-001-create-worksheet",
  "category": "工作表管理",
  "name": "创建工作表",
  "description": "创建一个名为 NewSheet 的工作表",
  "userInput": "创建一个名为 NewSheet 的工作表",
  "expectedCode": "Excel.run(async (context) => { ... })",
  "expectedBehavior": "新工作表 NewSheet 被创建并激活",
  "validationSteps": [
    "检查工作表列表是否包含 NewSheet",
    "检查 NewSheet 是否为活动工作表"
  ],
  "toolsTemplate": "工作表管理模板 > 创建工作表"
}
```

## 错误报告格式

错误报告包含以下信息:

```typescript
interface ErrorReport {
  timestamp: string;
  testCaseId: string;
  hostType: 'excel' | 'word' | 'powerpoint';
  errorType: 'InvalidArgument' | 'InvalidReference' | 'ApiNotFound' | 'GeneralException';
  errorCode?: string;
  errorMessage: string;
  stackTrace?: string;
  userInput: string;
  generatedCode: string;
  context: {
    officeVersion: string;
    platform: string;
  };
}
```

## 迭代优化流程

1. **执行测试** → 在 Office 环境中运行测试用例
2. **收集错误** → 自动记录所有错误详情
3. **分析错误** → 识别错误模式和高频问题
4. **优化 Skills** → 更新 SKILL.md 和 TOOLS.md
5. **更新测试** → 补充边界测试用例
6. **重复迭代** → 直到错误率 < 5%

## 注意事项

- 测试需要在真实的 Office 环境中运行
- 不同 Office 版本和平台可能有 API 差异
- 测试结果会自动保存到 `error-reports/` 目录
- 定期备份错误报告数据
