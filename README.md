# Vue Excel Import

> 基于 Vue 3 + TypeScript + Element Plus 的前端 Excel 导入组件。前端解析、校验、预览后分批提交到后端。

## 特性

### 基础功能

- 📤 **拖拽上传** — 支持拖拽和点击选择 .xlsx/.xls 文件
- 📋 **模板下载** — 根据列配置自动生成 Excel 模板（含必填标记、下拉验证、单元格格式）
- 🔄 **流式解析** — Web Worker 中使用 hucre 解析，按批发送到主线程，避免浏览器崩溃
- ⚙️ **列配置** — 动态配置列映射（Excel 表头 ⇢ 业务字段）、列类型、校验规则
- ✅ **6 种校验** — required / unique / range / pattern / enum / 自定义（支持异步）
- 📊 **数据预览** — 分页表格展示，错误单元格红色高亮 + hover 提示
- ✏️ **内联编辑** — 双击错误单元格直接修改，修改后自动重新校验
- 📦 **分批提交** — 默认 500 条/批，失败时自动回滚已提交批次
- 🔌 **可配置接口** — 支持字符串 URL 或自定义提交/回滚函数

### 高级功能

- 🧩 **多级表头** — `ColumnConfig.children` 支持嵌套分组表头，生成的模板自动合并单元格，上传时自动检测 Excel 合并单元格结构并匹配到配置
- 🏷️ **嵌套字段** — `field` 支持 `basic.name` 点号路径，数据自动组织为嵌套对象
- 🔍 **列匹配诊断** — 自动检测 Excel 中缺少的配置列和未匹配的多余列，界面中给出黄色提示
- 🎛️ **模板自定义** — 说明文字、示例数据行、Sheet 名称、冻结窗格

## 安装

```bash
npm install @scope/vue-excel-import
```

**peer 依赖**（需确保项目中已安装）：

```bash
npm install vue@^3.4 element-plus@^2.8
```

## 快速开始

```vue
<template>
  <ExcelImport
    :columns="columns"
    :submit-api="submitApi"
    :rollback-api="rollbackApi"
    @success="onSuccess"
    @error="onError"
  />
</template>

<script setup lang="ts">
import { ElNotification } from 'element-plus'
import ExcelImport from '@scope/vue-excel-import'
import type { ColumnConfig } from '@scope/vue-excel-import'

// 列配置（开发者初始化配置）
const columns: ColumnConfig[] = [
  { label: '姓名', field: 'userName', type: 'string', required: true },
  { label: '年龄', field: 'age', type: 'number', validators: [{ type: 'range', min: 0, max: 150 }] },
  { label: '邮箱', field: 'email', type: 'string', validators: [{ type: 'pattern', pattern: '^[\\w.-]+@[\\w.-]+\\.\\w+$' }] },
  { label: '部门', field: 'dept', type: 'string', validators: [{ type: 'enum', enum: ['技术部', '市场部', '财务部'] }] },
  { label: '入职日期', field: 'hireDate', type: 'date' },
]

// 提交接口（支持字符串 URL 或自定义函数）
const submitApi = '/api/user/batch-import'
const rollbackApi = '/api/user/rollback-import'

function onSuccess() { ElNotification({ title: '导入成功', type: 'success' }) }
function onError(err: Error) { ElNotification({ title: '导入失败', message: err.message, type: 'error' }) }
</script>
```

---

# 基础功能

## Props

| Prop | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| columns | `ColumnConfig[]` | — | **必填**。列配置 |
| submitApi | `string \| Function` | — | **必填**。提交接口 |
| rollbackApi | `string \| Function` | — | 回滚接口（分批失败时调用） |
| batchSize | `number` | 500 | 每批提交条数 |
| maxFileSize | `number` | 50 | 文件大小限制（MB） |
| showTemplateButton | `boolean` | true | 显示下载模板按钮 |
| templateFileName | `string` | "导入模板.xlsx" | 模板文件名 |
| templateSheetName | `string` | "Sheet1" | 模板 Sheet 名 |
| templateDescription | `string` | — | 模板顶部说明文字 |
| templateExampleRows | `number` | 0 | 模板中示例数据行数 |

## ColumnConfig

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| label | `string` | — | **必填**。Excel 表头名称（用于自动匹配列） |
| field | `string` | — | 映射到的业务字段名。支持 `basic.name` 形式的嵌套路径 |
| type | `'string' \| 'number' \| 'boolean' \| 'date'` | `'string'` | 解析类型 |
| required | `boolean` | false | 是否必填 |
| default | `any` | — | 空值时的默认值 |
| description | `string` | — | 预览时的字段说明 |
| validators | `ValidatorConfig[]` | — | 校验规则列表 |
| children | `ColumnConfig[]` | — | 子列配置（用于多级分组表头，详见高级功能） |

> `field` 和 `children` 互斥 — 有 `children` 的节点为分组节点，不绑定字段；叶子节点必须设置 `field`。

## ValidatorConfig

| type | 额外参数 | 说明 |
|------|----------|------|
| `required` | — | 必填校验（建议直接用 `ColumnConfig.required`） |
| `unique` | — | 整列值唯一 |
| `range` | `min`, `max` | 数字范围校验 |
| `pattern` | `pattern: string \| RegExp` | 正则表达式匹配 |
| `enum` | `enum: any[]` | 枚举值列表 |
| `custom` | `validate: (value, row, allData) => boolean \| string \| Promise` | 自定义校验函数 |

所有校验器可选 `message` 字段自定义错误提示。

### custom 校验函数示例

```typescript
{
  type: 'custom',
  message: '自定义校验未通过',
  validate: (value, row, allData) => {
    if (!value) return true
    return value.length >= 2 ? true : '长度不能小于 2'
  }
}

// 异步校验
{
  type: 'custom',
  validate: async (value) => {
    const res = await fetch(`/api/check?value=${value}`)
    return res.ok ? true : '值已被占用'
  }
}
```

## Events

| Event | 回调参数 | 触发时机 |
|-------|----------|----------|
| `@parsed` | `ParseResult` | 解析完成 |
| `@progress` | `SubmitProgress` | 分批提交进度更新 |
| `@success` | `SubmitResult` | 全部提交成功 |
| `@error` | `Error` | 提交失败 |

### ParseResult

| 字段 | 类型 | 说明 |
|------|------|------|
| `headers` | `string[]` | 解析到的 Excel 表头（叶子列名） |
| `rows` | `Record<string, any>[]` | 解析后的数据行 |
| `totalRows` | `number` | 总数据行数（不含表头） |
| `parseErrors` | `CellError[]` | 类型转换错误列表 |
| `sheets` | `string[]` | Excel 文件中的 Sheet 名称列表 |
| `missingColumns` | `string[]` | 配置了但在 Excel 中未匹配到的列名 |
| `unmatchedHeaders` | `string[]` | Excel 中存在但未在配置中找到匹配的列名 |

### SubmitProgress

| 字段 | 类型 | 说明 |
|------|------|------|
| `current` | `number` | 已提交条数 |
| `total` | `number` | 总条数 |
| `batch` | `number` | 当前批次号 |
| `totalBatches` | `number` | 总批次数 |

### SubmitResult

| 字段 | 类型 | 说明 |
|------|------|------|
| `success` | `boolean` | 是否全部提交成功 |
| `totalBatches` | `number` | 总批次数 |
| `completedBatches` | `number` | 成功提交批次数 |
| `failedBatch` | `number` | 失败批次号（失败时） |
| `rollbackSuccess` | `boolean` | 回滚是否成功 |
| `error` | `string` | 错误信息（失败时） |

## 类型转换规则

组件根据 `column.type` 自动转换 Excel 原始值，转换失败不中断解析，错误收集到 `parseErrors`。

**`type: 'string'`**
- number → "25"（自动转字符串）
- boolean → "true" / "false"
- Date → "2024-01-15"（格式化为 YYYY-MM-DD）
- 自动 trim 前后空格

**`type: 'number'`**
- 字符串 "42" → 42（自动解析）
- boolean → 1 / 0
- 非法字符串 → 报 parseError

**`type: 'boolean'`**
- 0 → false, 非 0 → true
- "true" / "1" / "yes" → true
- "false" / "0" / "no" → false

**`type: 'date'`**
- Date 对象 → 直接格式化
- 数字 > 100000 → Unix 毫秒时间戳
- 数字 ≤ 100000 → Excel 序列号（含闰年 bug 补偿）
- 字符串 → new Date() 解析

## 模板下载

组件可根据 `columns` 配置自动生成 Excel 模板：

- **必填标记** — 必填列的列头标注 `*`（如 `姓名*`），上传时 `*` 会被自动忽略不影响匹配
- **单元格格式** — number 列设数字格式、date 列设日期格式
- **下拉验证** — enum 类型列自动添加 Excel 数据验证（下拉选择）
- **冻结窗格** — 自动冻结表头行，方便滚动查看

## 分批提交

数据通过校验后，按 `batchSize`（默认 500）分批顺序提交：

- 每批提交后更新进度（`@progress`）
- 某批失败时自动调用 `rollbackApi` 回滚已提交批次
- 支持字符串 URL（用 fetch POST JSON）或自定义函数

---

# 高级功能

## 嵌套字段路径

`field` 支持用 `.` 分隔的多级路径，数据会被自动组织为嵌套对象：

```typescript
const columns: ColumnConfig[] = [
  { label: '姓名', field: 'basic.name', type: 'string', required: true },
  { label: '年龄', field: 'basic.age', type: 'number' },
  { label: '部门', field: 'dept', type: 'string' }
]
```

解析后每行数据格式：

```json
{
  "basic": { "name": "张三", "age": 28 },
  "dept": "技术部"
}
```

嵌套字段可通过 `default` 设置默认值，或通过 `validators` 添加校验规则。

## 多级表头（分组列）

通过 `ColumnConfig.children` 定义嵌套分组表头，支持任意层级：

```typescript
const columns: ColumnConfig[] = [
  {
    label: '基本信息',
    children: [
      { label: '姓名', field: 'name', type: 'string', required: true },
      { label: '年龄', field: 'age', type: 'number' },
      {
        label: '联系方式',
        children: [
          { label: '邮箱', field: 'email', type: 'string' },
          { label: '电话', field: 'phone', type: 'string' },
        ]
      },
    ]
  },
  { label: '部门', field: 'dept', type: 'string' }
]
```

- **模板生成** — 分组列自动合并单元格，支持任意层级嵌套
- **自动检测** — 上传时自动识别 Excel 的合并单元格结构，通过表头标签匹配到配置列
- **兼容性** — 单级表头（无 `children`）完全兼容，无需修改现有配置

## 列匹配诊断

解析完成后，组件自动比对配置列和 Excel 实际列：

- **缺少列** → 黄色警告："Excel 中缺少配置列: 姓名、邮箱"。这些列的数据无法导入
- **多余列** → 蓝色提示："Excel 中存在未配置的列: 备注"。这些列将被忽略

诊断信息同时通过 `@parsed` 事件的 `ParseResult.missingColumns` 和 `ParseResult.unmatchedHeaders` 暴露给开发者。

### 列匹配流程

1. 读取 Excel 原始行数据和合并单元格信息
2. `detectHeaderTree()` 根据合并单元格和行结构重建表头树
3. `matchColumnTree()` 按 `label` 匹配配置树和 Excel 树，自动忽略 `*` 标记
4. 生成 `columnMap`（列索引 → 配置叶节点）和诊断信息

## 模板自定义

通过以下 props 控制模板生成行为：

| Prop | 效果 |
|------|------|
| `template-description` | 模板顶部添加说明行（如"请按格式填写，标 * 为必填"） |
| `template-example-rows` | 生成 N 行示例数据，按类型填充（number=0, date=当天, boolean=true） |
| `template-file-name` | 下载时的文件名 |
| `template-sheet-name` | Sheet 名称 |
| `show-template-button` | 是否显示下载按钮 |

## Web Worker 解析

解析在 Web Worker 中执行，避免阻塞主线程：

- **Worker 可用** → `excelParser.worker.ts` 在后台线程用 hucre 读取，按 500 行为一批发送到主线程
- **Worker 不可用** → 自动降级到主线程解析（同批处理逻辑），支持非标准浏览器环境
- **进度反馈** → `parser.progress` 响应式变量从 0 更新到 100，组件内自动显示进度条

---

# 开发

## 本地开发

```bash
# 克隆项目
git clone <repo-url>
cd vue-excel-import

# 安装依赖
npm install

# 运行 demo（dev 模式下以 demo/ 为入口启动）
npm run dev

# 运行测试
npm test

# 构建
npm run build
```

## 构建产物

```bash
npm run build
# dist/
# ├── vue-excel-import.js          # ESM
# ├── vue-excel-import.umd.cjs     # UMD
# ├── style.css                    # 样式
# ├── index.d.ts                   # 类型声明
# ├── types/                       # 类型定义
# ├── composables/
# ├── components/
# └── utils/
```

## License

MIT
