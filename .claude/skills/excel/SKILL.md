---
name: excel-operations
description: Excel 数据操作技能。用于单元格读写、工作表管理、表格与数据透视表创建、数据验证、批注、图表及数据分析。当用户提及 Excel、表格、单元格、图表、数据透视表或数据分析时使用。
---

# Excel 操作技能

## 用法

通过生成**隐藏的 Office.js 代码**来操作 Excel，这些代码由前端自动执行，对用户完全透明。

### 重要规则

- **用户体验优先**：用户只需看到自然语言，不应看到任何代码
- **隐藏代码格式**：用 HTML 注释包裹代码：`<!--OFFICE-CODE:excel\ncode\n-->`
- **友好反馈**：操作完成后用自然语言告知用户结果
- **完整且可执行**：生成的代码必须是完整、可直接运行的 Office.js 代码
- **读取操作需返回数据**：当用户要求**读取**、**获取**或**查看**区域/工作表数据时，生成的代码**必须**在 `Excel.run` 回调内 `return` 该数据（如 `return range.values`、`return { address, values, formulas }`）。执行器会将返回值传回 AI 以供分析或回答——没有 `return` 时，AI 无法接收数据。参见 TOOLS.md 中的「数据读取模板」（读取选中区域、读取指定区域）；这些模板中的 `return` 是必需的，不得省略。

## ⚠️ API 使用指南与常见错误模式

### 理解 Office.js API 结构

Office.js Excel API 遵循一致的模式。理解这些模式可避免大多数常见错误：

**1. 集合-项模式**
- 集合始终为复数：`worksheets`、`tables`、`charts`、`borders`、`pivotTables`
- 项为单数，通过 `getItem(name/id)`、`getItemAt(index)`、`getItemOrNullObject(name/id)` 访问
- 规则：切勿在项上调用集合方法，或在集合上调用项方法

**2. 格式对象层级**
- 格式属性嵌套：`range.format.{category}.{property}`
- 类别：`font`、`fill`、`borders`、`protection`、对齐属性
- 边框操作需使用 **`borders`** 集合（复数），而非 `border`（单数）

**3. 基于数组的属性**
- `values`、`formulas`、`numberFormat` 等属性始终期望二维数组
- 规则：数组维度必须与区域维度完全匹配
- 即使单个单元格也需要二维数组结构：`[[value]]`

### 多区域操作（RangeAreas）

- ❌ **错误**：`sheet.getRangeAreas("B3,F3,J3")` - 工作表对象没有 getRangeAreas 方法
- ✅ **正确**：`workbook.getRangeAreas("Sheet1!B3,Sheet1!F3,Sheet1!J3")` - 必须使用 Workbook 对象并包含工作表名称
- ✅ **推荐**：分别操作每个单元格：`sheet.getRange("B3")`、`sheet.getRange("F3")` 等

**何时使用 RangeAreas：**
- 跨工作表操作
- 非连续区域的批量格式设置
- 相比单个 Range 对象支持的操作有限

**何时使用单独的 Range 调用：**
- 为每个区域设置不同值
- 对每个区域进行复杂操作
- 更易读、易维护的代码

### 错误模式与预防

**模式 1：InvalidArgument - 无效或缺失参数**
- **原因**：不存在的对象引用、参数类型错误、值超出范围
- **预防**：操作前使用 `getItemOrNullObject()` 并检查 `isNullObject`
```javascript
const sheet = context.workbook.worksheets.getItemOrNullObject("SheetName");
await context.sync();
if (sheet.isNullObject) {
  console.log("Worksheet doesn't exist");
  return;
}
```

**模式 2：InvalidReference - 无效的单元格/区域引用**
- **原因**：引用的单元格或区域在工作表中不存在
- **预防**：访问前验证区域地址格式和边界
```javascript
const range = sheet.getRange("A1:Z1000");
range.load("address");
await context.sync();
// Range is now validated and safe to use
```

**模式 3：数组维度不匹配**
- **错误信息**："The number of rows or columns in the input array doesn't match the size or dimensions of the range"
- **原因**：设置 `values`、`formulas` 或 `numberFormat` 时数组大小与区域大小不匹配
- **预防**：根据区域大小计算数组维度
```javascript
// For range "A1:B2" (2 rows × 2 columns)
range.values = [[1, 2], [3, 4]]; // ✅ Correct: 2×2 array

// For single cell "A1" (1 row × 1 column)
range.values = [[100]]; // ✅ Correct: 1×1 array
```

**模式 4：属性/方法未找到**
- **错误信息**："undefined is not an object" 或 "is not a function"
- **原因**： 
  - 使用单数而非复数（如 `border` 而非 `borders`）
  - 在项上调用集合方法或反之
  - 在 `load()` 和 `sync()` 之前访问属性
- **预防**：遵循 API 层级，使用正确的集合名称，访问前先 load 属性

**模式 5：请求的资源不存在**
- **原因**：尝试删除或修改不存在的对象
- **预防**：删除操作前始终使用 `getItemOrNullObject()`
```javascript
const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
await context.sync();
if (!pivotTable.isNullObject) {
  pivotTable.delete();
  await context.sync();
}
```

### Load 与 Sync 最佳实践

1. **Load 模式**：读取前必须先 load 属性。当用户要求**读取**供 AI 使用的数据时，需在 `context.sync()` 后 **return** 该数据（如 `return range.values` 或 `return { address, values }`）——参见 TOOLS.md 数据读取模板。
   ```javascript
   range.load("values, address, formulas");
   await context.sync();
   return { address: range.address, values: range.values, formulas: range.formulas }; // 读取操作必需
   ```

2. **Write 模式**：设置属性后 sync 以提交更改
   ```javascript
   range.values = [[100]];
   range.format.font.bold = true;
   await context.sync(); // 提交所有更改
   ```

3. **批量操作**：尽量少调 `sync()` 以提升性能
   ```javascript
   // ✅ 好：多次操作一次 sync
   range1.values = [[1]];
   range2.values = [[2]];
   range3.values = [[3]];
   await context.sync();
   
   // ❌ 差：多次 sync
   range1.values = [[1]];
   await context.sync();
   range2.values = [[2]];
   await context.sync();
   ```

4. **跨上下文访问**：在一个上下文中 load 的属性在另一个上下文中无法访问
   ```javascript
   // ❌ 错误：尝试在 Excel.run 外使用 load 的数据
   let data;
   await Excel.run(async (context) => {
     const range = sheet.getRange("A1");
     range.load("values");
     await context.sync();
     data = range.values; // 在上下文结束前存储
   });
   console.log(data); // ✅ 现在可在上下文外访问
   ```

### 异步操作与上下文管理

1. **上下文作用域**：每个 `Excel.run()` 创建独立上下文
   - 在一个上下文中创建的对象无法在另一个上下文中直接使用
   - 存储原始值（字符串、数字、数组）以在上下文间传递

2. **异常处理**：用 try-catch 包裹操作以稳健处理错误
   ```javascript
   try {
     await Excel.run(async (context) => {
       // Operations here
       await context.sync();
     });
   } catch (error) {
     console.error("Excel operation failed:", error);
     // Provide user-friendly error message
   }
   ```

3. **Async/Await 要求**：
   - 所有 `context.sync()` 调用必须 await
   - Excel.run 回调必须是 async 函数
   - 不要在没有 proper async/await 处理的情况下使用 callbacks 或 promises

### 枚举常量与类型安全

始终使用 Excel 命名空间枚举以提升代码质量：

```javascript
// ✅ 正确：使用枚举
range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
border.style = Excel.BorderLineStyle.continuous;
border.weight = Excel.BorderWeight.thick;

// ❌ 避免：使用字符串（易出错）
range.format.horizontalAlignment = "center"; // May work but not type-safe
border.style = "Continuous"; // Case-sensitive, easy to typo
```

常用枚举命名空间：
- `Excel.BorderIndex`: edgeTop, edgeBottom, edgeLeft, edgeRight, insideHorizontal, insideVertical
- `Excel.BorderLineStyle`: none, continuous, dash, dashDot, dot, double
- `Excel.BorderWeight`: hairline, thin, medium, thick
- `Excel.HorizontalAlignment`: left, center, right, justify, distributed
- `Excel.VerticalAlignment`: top, center, bottom, justify, distributed
- `Excel.ChartType`: columnClustered, line, pie, bar, area, scatter, etc.
- `Excel.RangeUnderlineStyle`: none, single, double, singleAccountant, doubleAccountant

### 参数验证规则

1. **区域地址**： 
   - 必须遵循 A1 表示法："A1"、"B2:D10"、"Sheet1!A1:B5"
   - 列字母大小写不敏感："A1" 等于 "a1"
   - 行号从 1 开始，非 0

2. **数组数据**：
   - 外层数组表示行，内层数组表示列
   - 所有内层数组长度必须相同
   - 空单元格应为空字符串 "" 或 null，不应为 undefined

3. **对象名称**：
   - 工作表名称：避免特殊字符 `[]/*?:`
   - 表格名称：必须以字母或下划线开头，不能有空格
   - 命名区域：类似表格名称，遵循 Excel 命名规则

### 性能优化指南

1. **最小化 Sync 调用**：每次 sync 都是与 Excel 的往返
   ```javascript
   // ✅ 高效：100 次操作 1 次 sync
   for (let i = 0; i < 100; i++) {
     sheet.getRange(`A${i+1}`).values = [[i]];
   }
   await context.sync();
   ```

2. **使用批量操作**：优先使用区域操作而非逐单元格
   ```javascript
   // ✅ 更好：单次区域操作
   range.values = [[1,2,3], [4,5,6], [7,8,9]];
   
   // ❌ 更慢：多次单元格操作
   sheet.getRange("A1").values = [[1]];
   sheet.getRange("B1").values = [[2]];
   // ... 7 more operations
   ```

3. **仅 Load 所需属性**：减少数据传输
   ```javascript
   // ✅ 高效：Load 特定属性
   range.load("values, address");
   
   // ❌ 低效：Load 全部
   range.load();
   ```

4. **批量相似操作**：将相似操作分组
   ```javascript
   // 先设置所有格式，再设置所有值
   range1.format.font.bold = true;
   range2.format.font.bold = true;
   range3.format.font.bold = true;
   range1.values = [[1]];
   range2.values = [[2]];
   range3.values = [[3]];
   await context.sync();
   ```

### 功能特定 API 指南

#### 条件格式
- 每个区域可应用多个条件格式，按优先级顺序评估
- 使用 `conditionalFormat.priority` 调整评估顺序
- 使用 `conditionalFormat.stopIfTrue` 阻止后续规则评估
- 图标集 `criteria` 数组：索引 0 为最低级，最后索引为最高级
- 格式选项通过嵌套对象访问：`conditionalFormat.{type}.format.{category}`

#### 事件处理
- 加载项刷新或关闭时事件处理器会被销毁
- 模式：保存 `EventResult` 对象 → 调用 `eventResult.remove()` 清理
- 批量操作期间可临时禁用事件：`context.runtime.enableEvents = false`
- 事件为工作表或工作簿级别，通过 `sheet.on{EventName}.add(handler)` 访问

#### 形状操作
- 位置属性（`left`、`top`）使用磅（points），非像素
- 图片插入需要 Base64 字符串，不含 `data:image/png;base64,` 前缀
- 形状分组：先通过 `shape.load("id")` 加载形状 `id` 属性
- 命名形状以便后续引用：`shape.name = "MyShape"`
- 形状集合通过 `sheet.shapes` 或 `chart.shapes` 访问

#### 批注 vs 评论
- **批注（Notes）**：传统黄色便签（每单元格一个）
- **评论（Comments）**：线程式讨论评论（每单元格多个）
- 访问方式：`worksheet.notes` 或 `workbook.notes` 获取批注集合
- 属性 `note.visible` 控制永久可见性（默认：仅悬停时显示）
- 使用 `notes.add(cellAddress, content)` 创建批注

#### 区域高级操作
- `copyFrom(source, copyType, skipBlanks, transpose)`：skipBlanks 保留目标数据
- `moveTo(destinationRange)`：剪切粘贴操作，自动扩展目标
- `removeDuplicates(columns, includesHeader)`：列索引从 0 开始
- `insert(shift)` 和 `delete(shift)`：影响工作表中的周围区域
- 使用 insert/delete 时始终考虑对其他区域的影响

#### 复选框
- 用于布尔值可视化的单元格控件类型
- 转换：`range.control = { type: Excel.CellControlType.checkbox }`
- 状态管理：使用 `range.values` 配合 `[[true]]` 或 `[[false]]`
- 读取状态：`range.load("values")` 后 `range.values` 返回布尔数组
- 移除：`range.control = { type: Excel.CellControlType.empty }`

## ⚠️ 工具选择优先级（强制规则）

### 优先使用 MCP 领域工具

DocuPilot 2.0 提供**领域聚合 MCP 工具**，比通用 execute_code 工具更快、更安全、更易用。

**强制规则**：
1. **默认使用 MCP 领域工具** - 覆盖 85%+ 常见场景
2. **仅在 MCP 工具无法满足需求时使用 execute_code** - 用于复杂高级 API

### 可用 Excel MCP 工具

| 工具 | 用途 | 频率 |
|------|---------|-----------|
| `excel_range` | 单元格读写、格式设置 | ⭐⭐⭐ 最常用 |
| `excel_worksheet` | 工作表管理 | ⭐⭐ 常用 |
| `excel_table` | 表格对象操作 | ⭐ 中等 |
| `excel_chart` | 图表创建与管理 | ⭐ 中等 |
| `execute_code` | 高级 API（数据透视表等） | 备用工具 |

### 工具选择决策树

```
用户请求
  |
  ├─ 读取/写入单元格？ → 使用 excel_range
  ├─ 管理工作表？ → 使用 excel_worksheet  
  ├─ 创建/操作表格？ → 使用 excel_table
  ├─ 创建图表？ → 使用 excel_chart
  └─ 数据透视表/条件格式/复杂批量？ → 使用 execute_code
```

### MCP 工具调用方式

```typescript
// ✅ 推荐：使用 MCP 领域工具
mcp__office__excel_range({
  action: "read",
  address: "A1:C10",
  includeFormulas: true
})

// ❌ 不推荐：除非 MCP 工具无法满足需求
mcp__office__execute_code({
  host: "excel",
  code: "Excel.run(async (context) => { ... })"
})
```

### 示例对比

**场景**：从 Sheet1 读取财务数据

**使用 MCP 工具（推荐）**：
```typescript
// 步骤 1：读取数据
mcp__office__excel_range({
  action: "read",
  address: "Sheet1!A1:F100",
  includeFormulas: false
})

// 步骤 2：如需则格式化
mcp__office__excel_range({
  action: "format",
  address: "A1:F1",
  format: {
    font: { bold: true, size: 14 },
    fill: { color: "#4472C4" }
  }
})
```

**使用 execute_code（仅必要时）**：
```typescript
// 仅当需要高级 API（如数据透视表）时
mcp__office__execute_code({
  host: "excel",
  description: "Create PivotTable to analyze financial data",
  code: `
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      const pivotTable = sheet.pivotTables.add("FinancePivot", range, "G1");
      // ... configure pivot table
      await context.sync();
      return { success: true };
    });
  `
})
```

### 性能对比

| 指标 | MCP 工具 | execute_code | 改进 |
|--------|-----------|--------------|-------------|
| 响应时间 | 1.2s | 2.5s | ↓52% |
| Token 消耗 | ~280 | ~800 | ↓65% |
| 错误率 | <5% | 15% | ↓67% |

### 完整工具 API 参考

详细工具参数和返回值请参考：
- [MCP工具API文档](../../../docs/MCP_TOOLS_API.md)
- [MCP工具完整列表](../../../docs/MCP_TOOLS_REFERENCE.md)

## 工作流

1. **理解需求**：分析用户的数据操作请求
2. **参考模板**：查阅 TOOLS.md 中的代码模板
3. **生成代码**：创建完整的 Office.js 代码
4. **嵌入隐藏标记**：用 `<!--OFFICE-CODE:excel ... -->` 包裹代码
5. **添加友好消息**：告知用户操作结果

## 支持的功能

- **工作表管理**：创建、重命名、删除、复制、激活工作表。
  - ⚠️ 删除操作前必须使用 `getItemOrNullObject` 检查对象存在性
- **数据读写**：读写单元格、区域、数组数据。
- **表格操作**：创建表格、排序、筛选、添加行/列。
- **数据透视表**（完整支持）：
  - 创建数据透视表，添加行/列/数据/筛选字段
  - 应用筛选（日期筛选、标签筛选、值筛选）
  - 创建和使用切片器进行交互式筛选
  - 切换布局类型（紧凑、大纲、表格）
  - 获取和格式化数据透视表数据
  - 刷新和删除数据透视表
  - ⚠️ 所有操作前必须检查数据透视表存在性
- **条件格式**（完整支持）：
  - 单元格值条件（cellValue）：根据单元格值应用格式
  - 数据条（dataBar）：在单元格中显示数据条
  - 图标集（iconSet）：用箭头、图标等可视化数据
  - 预设条件（preset）：高于/低于平均值等预设规则
  - 前 N/后 N（topBottom）：高亮前 N 或后 N
  - 自定义公式（custom）：使用自定义公式设置条件
- **事件处理**：
  - 数据变化事件（onChanged）：监控单元格数据变化
  - 选择变化事件（onSelectionChanged）：监控用户选择变化
  - 工作表激活事件（onActivated）：监控工作表切换
  - 计算完成事件（onCalculated）：监控工作表计算完成
  - ⚠️ 事件处理器需正确注册和移除以避免内存泄漏
- **形状操作**：
  - 几何形状：添加矩形、圆形、箭头等
  - 图片插入：支持 Base64 编码的 JPEG/PNG 图片
  - 线条和连接器：创建直线、箭头线等
  - 文本框：创建和格式化文本框
  - 形状分组：将多个形状组合成组
- **批注**：
  - 添加单元格批注（传统黄色便签）
  - 修改批注内容和可见性
  - 删除批注
  - ⚠️ 不同于评论（线程评论），批注是传统批注功能
- **区域高级操作**：
  - 复制粘贴：支持跳过空、转置选项
  - 移动区域：剪切并粘贴到新位置
  - 插入/删除单元格：指定移位方向
  - 清除内容：清除值、格式或全部
  - 删除重复项：根据指定列删除重复行
  - 行/列分组：创建大纲组并取消分组
- **复选框**：
  - 将布尔值转换为复选框控件
  - 读取复选框状态
  - 移除复选框恢复为布尔值
- **数据验证**：设置数值范围、日期限制、下拉列表等验证规则。
- **评论与命名**：添加评论、创建命名区域。
- **图表**：创建各类图表（柱状图、折线图等）。
- **格式设置**：设置数字格式、字体、颜色。

## 数据透视表操作注意事项

### 前置检查
数据透视表操作前必须检查对象存在性，避免「请求的资源不存在」错误：

```javascript
const pivotTable = sheet.pivotTables.getItemOrNullObject("PivotTable1");
await context.sync();

if (pivotTable.isNullObject) {
  console.log("Pivot table doesn't exist");
  return;
}
```

### 字段引用
数据透视表字段必须来自源数据列名，确保：
1. 字段名与源数据列标题完全匹配
2. 使用 `getItemOrNullObject` 检查字段存在性
3. 操作字段前先 sync 确保字段已加载

### 筛选操作
- **日期筛选**：需指定日期条件和精度（年/月/日）
- **标签筛选**：针对文本字段，支持包含/排除、以/结尾于等条件
- **值筛选**：按聚合值筛选，需指定数值字段和比较条件

### 切片器创建
切片器依赖数据透视表，创建前必须：
1. 确保数据透视表存在
2. 确保筛选字段在数据透视表中存在
3. 设置合理的位置和大小以避免重叠

## 示例

### 示例 1：管理工作表
**用户请求**：「创建一个名为「销售报告」的新工作表」

**你的回复**：
```
好的，我来为你创建新工作表。

[调用 MCP 工具：excel_worksheet，action "add"]

✅ 已成功创建并激活工作表「销售报告」。
```

**使用的工具**：`mcp__office__excel_worksheet({ action: "add", name: "Sales Report", position: "end" })`

### 示例 2：读取并分析数据
**用户请求**：「读取 A1:D10 的数据并格式化表头行」

**你的回复**：
```
我来读取数据并为你格式化。

[步骤 1：使用 excel_range 工具读取数据]
[步骤 2：使用 excel_range 工具格式化表头行]

✅ 已成功读取 10 行数据，并将表头行格式化为粗体、蓝色背景。
```

**使用的工具**： 
1. `mcp__office__excel_range({ action: "read", address: "A1:D10", includeFormulas: false })`
2. `mcp__office__excel_range({ action: "format", address: "A1:D1", format: { font: { bold: true }, fill: { color: "#4472C4" } } })`

### 示例 3：创建数据透视表（高级 API）
**用户请求**：「用当前数据创建数据透视表，按产品汇总销售」

**你的回复**：
```
我来帮你创建数据透视表。这需要高级 API 功能。

<!--OFFICE-CODE:excel
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getUsedRange();
  
  const pivotTable = sheet.pivotTables.add("PivotTable1", range, sheet.getRange("F1"));
  
  pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Product"));
  pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Sales"));
  
  await context.sync();
});
-->

✅ 已创建数据透视表，按产品汇总销售数据，起始位置为 F1 单元格。
```

**说明**：数据透视表需使用 execute_code，因其使用的高级 API 尚未被 MCP 领域工具覆盖。
### 示例 3：设置数据验证（高级 API）
**用户请求**：「将 B 列限制为仅接受 1 到 100 之间的整数」

**你的回复**：
```
我来为 B 列设置数据验证规则。这需要高级 API 功能。

<!--OFFICE-CODE:excel
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("B:B");
  
  range.dataValidation.rule = {
    wholeNumber: {
      formula1: 1,
      formula2: 100,
      operator: Excel.DataValidationOperator.between
    }
  };
  
  range.dataValidation.errorAlert = {
    showAlert: true,
    title: "Input Error",
    message: "Please enter an integer between 1 and 100",
    style: Excel.DataValidationAlertStyle.stop
  };
  
  await context.sync();
});
-->

✅ 已为 B 列设置数据验证：仅允许 1-100 的整数。
```

**说明**：数据验证需使用 execute_code，因其使用的高级 API 尚未被 MCP 领域工具覆盖。

## 用户文件处理

### 工作区文件访问

用户可将文件上传到当前会话，这些文件保存在工作区目录：

- **用户上传文件路径**：`workspace/sessions/{session_id}/uploads/`
- **生成文件保存路径**：`workspace/sessions/{session_id}/outputs/`

### 文件操作流程

1. **查找用户上传的文件**：
   ```typescript
   // 使用 Glob 工具查找 Excel 文件
   // 文件名包含时间戳前缀，使用通配符
   const pattern = "workspace/sessions/{session_id}/uploads/*.xlsx";
   ```

2. **读取文件数据**：
   - 文本格式文件（CSV、TXT、JSON）使用 Read 工具
   - Excel 文件引导用户用 Excel 打开后使用 Office.js API 操作

3. **保存分析结果**：
   ```typescript
   // 使用 Write 工具保存分析报告
   Write: workspace/sessions/{session_id}/outputs/analysis_report.txt
   ```

### 示例工作流

**用户请求**：「分析我上传的销售数据表」

**处理步骤**：
1. 使用 Glob 查找：`workspace/sessions/abc123/uploads/*.xlsx`
2. 引导用户：「我找到你上传的文件 `sales_data.xlsx`。请用 Excel 打开此文件，然后我可以帮你分析数据。」
3. 用户用 Excel 打开文件后，使用 `office_excel_*` 工具读取和分析数据
4. 将分析结果保存到：`workspace/sessions/abc123/outputs/sales_analysis.txt`

## 详细模板

更多操作模板请参考 [TOOLS.md](TOOLS.md)。
