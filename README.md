# Vue Excel Import

> 基于 Vue 3 + TypeScript + Element Plus 的前端 Excel 导入组件。前端解析、校验、预览后分批提交到后端。

## 特性

- 📤 **拖拽上传** — 支持拖拽和点击选择 .xlsx/.xls 文件
- 📋 **模板下载** — 根据列配置自动生成 Excel 模板（含下拉验证、必填标记）
- 🔄 **流式解析** — 使用 hucre 解析，只读取有内容行，避免浏览器崩溃
- ⚙️ **列配置** — 动态配置列映射（Excel 表头 ⇢ 业务字段）、列类型、列选择
- ✅ **6 种校验** — required / unique / range / pattern / enum / 自定义（支持异步）
- 📊 **数据预览** — 分页表格展示，错误单元格红色高亮 + hover 提示
- ✏️ **内联编辑** — 双击错误单元格直接修改，修改后自动重新校验
- 📦 **分批提交** — 默认 500 条/批，支持失败回滚
- 🔌 **可配置接口** — 支持字符串 URL 或自定义提交/回滚函数

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

### ColumnConfig

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| label | `string` | — | **必填**。Excel 表头名称（用于自动匹配列） |
| field | `string` | — | **必填**。映射到的业务字段名 |
| type | `'string' \| 'number' \| 'boolean' \| 'date'` | `'string'` | 解析类型 |
| required | `boolean` | false | 是否必填 |
| default | `any` | — | 空值时的默认值 |
| description | `string` | — | 预览时的字段说明 |
| validators | `ValidatorConfig[]` | — | 校验规则列表 |

### ValidatorConfig

| type | 额外参数 | 说明 |
|------|----------|------|
| `required` | — | 必填校验（建议直接用 `ColumnConfig.required`） |
| `unique` | — | 整列值唯一 |
| `range` | `min`, `max` | 数字范围校验 |
| `pattern` | `pattern: string \| RegExp` | 正则表达式匹配 |
| `enum` | `enum: any[]` | 枚举值列表 |
| `custom` | `validate: (value, row, allData) => boolean \| string \| Promise` | 自定义校验函数 |

所有校验器可选 `message` 字段自定义错误提示。

#### custom 校验函数示例

```typescript
{
  type: 'custom',
  message: '自定义校验未通过',
  validate: (value, row, allData) => {
    if (!value) return true
    // 返回 true = 通过，返回字符串 = 错误提示
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

- 必填列的列头标注 `*`（如 `姓名*`）
- number/date 类型列自动设置单元格格式
- enum 类型列自动添加 Excel 数据验证（下拉选择）
- 可选生成示例数据行
- 上传时表头的 `*` 会被自动忽略，不影响列匹配

## 分批提交

数据通过校验后，按 `batchSize`（默认 500）分批顺序提交：

- 每批提交后更新进度（`@progress`）
- 某批失败时自动调用 `rollbackApi` 回滚已提交批次
- 支持字符串 URL（用 fetch POST JSON）或自定义函数

## 本地开发

```bash
# 克隆项目
git clone <repo-url>
cd vue-excel-import

# 安装依赖
npm install

# 运行 demo
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
