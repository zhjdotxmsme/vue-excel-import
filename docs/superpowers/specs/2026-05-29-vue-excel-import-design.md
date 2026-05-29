# Vue Excel Import Component — 设计文档

> 基于 Vue 3 + TypeScript + Element Plus 的前端 Excel 导入组件，前端解析、校验、预览后分批提交到后端。

## 1. 概述

封装一个通用的 Excel 导入组件，支持开发者通过 Props 配置列映射和校验规则，上传者上传 Excel → 自动解析 → 预览数据 → 修正错误 → 分批提交到后端。组件发布为 npm 私库包。

## 2. 使用示例

```vue
<template>
  <ExcelImport
    :columns="columns"
    :submit-api="submitApi"
    :rollback-api="rollbackApi"
    :batch-size="500"
    @success="onSuccess"
    @error="onError"
  />
</template>

<script setup lang="ts">
import ExcelImport from '@scope/vue-excel-import'

const columns = [
  { label: '姓名', field: 'userName', type: 'string', required: true },
  { label: '年龄', field: 'age', type: 'number', validators: [{ type: 'range', min: 0, max: 150 }] },
  { label: '邮箱', field: 'email', type: 'string', validators: [{ type: 'pattern', pattern: '^[\\w.-]+@[\\w.-]+\\.\\w+$' }] },
  { label: '部门', field: 'dept', type: 'string' },
  { label: '入职日期', field: 'hireDate', type: 'date' },
]
</script>
```

## 3. Props 设计

| Prop | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| columns | `ColumnConfig[]` | — | **必填**。列配置，定义 Excel 表头到业务字段的映射关系 |
| submitApi | `string \| Function` | — | **必填**。提交接口地址或自定义提交函数 |
| rollbackApi | `string \| Function` | — | 分批失败时的回滚接口（可选） |
| batchSize | `number` | 500 | 每批提交条数 |
| maxFileSize | `number` | 50 | 文件大小限制（MB） |
| showTemplateButton | `boolean` | true | 是否显示下载模板按钮 |
| templateFileName | `string` | "导入模板.xlsx" | 模板文件名 |
| templateSheetName | `string` | "Sheet1" | 模板 Sheet 名 |
| templateDescription | `string` | — | 模板顶部说明文字 |
| templateExampleRows | `number` | 0 | 模板中生成示例数据行数 |

### ColumnConfig

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| label | `string` | — | **必填**。Excel 表头名称，用于自动匹配列 |
| field | `string` | — | **必填**。映射到的业务字段名 |
| type | `'string' \| 'number' \| 'boolean' \| 'date'` | `'string'` | 解析类型 |
| required | `boolean` | false | 是否必填 |
| default | `any` | — | 空值时的默认值 |
| description | `string` | — | 预览时的字段说明 |
| validators | `ValidatorConfig[]` | — | 校验规则列表 |

### ValidatorConfig

| 字段 | 类型 | 适用 type | 说明 |
|------|------|-----------|------|
| type | `'required' \| 'unique' \| 'range' \| 'pattern' \| 'enum' \| 'custom'` | 全部 | 校验规则类型 |
| min | `number` | range | 最小值 |
| max | `number` | range | 最大值 |
| pattern | `string \| RegExp` | pattern | 正则表达式 |
| enum | `any[]` | enum | 允许的值列表 |
| message | `string` | 全部 | 自定义错误提示 |
| validate | `Function` | custom | 自定义校验函数 `(value, row, allData) => boolean \| string \| Promise` |

## 4. Events

| Event | 回调参数 | 触发时机 |
|-------|----------|----------|
| `@parsed` | `ParseResult` | 文件解析完成 |
| `@progress` | `SubmitProgress` | 分批提交进度更新 |
| `@success` | `SubmitResult` | 全部提交成功 |
| `@error` | `Error` | 提交失败（含回滚结果） |

## 5. 模板下载

根据 columns 配置自动生成 Excel 模板文件（.xlsx），通过 `useExcelTemplate` composable 实现：

- 顶部写入说明文字（可配置）
- 表头行写入 column.label，必填列追加 `*` 标记
- 根据 type 设置列格式（数字列设为数字格式，日期列设为 yyyy-mm-dd）
- enum 类型的列添加 Excel 数据验证（下拉列表）
- 可选生成示例数据行引导填写
- 通过 Blob + URL.createObjectURL + `<a>.click()` 触发下载

## 6. 内部架构

### 目录结构

```
src/
├── index.ts                   # 入口导出
├── components/
│   ├── ExcelImport.vue        # 主组件，编排子组件生命周期
│   ├── ExcelUploader.vue      # 上传区（拖拽/点击 + 模板下载按钮）
│   ├── ExcelPreview.vue       # 数据预览表格（分页 + 错误高亮）
│   ├── ExcelCellEdit.vue      # 错误单元格内联编辑
│   └── SubmitBar.vue          # 底部操作栏（统计 + 提交按钮 + 进度条）
├── composables/
│   ├── useExcelParser.ts      # ExcelJS 流式解析
│   ├── useExcelValidator.ts   # 校验引擎
│   ├── useExcelSubmitter.ts   # 分批提交 + 回滚
│   └── useExcelTemplate.ts    # 模板生成下载
├── types/
│   └── index.ts               # 所有类型定义
└── utils/
    └── excel.ts               # ExcelJS 通用工具函数
```

### 数据流

```
上传文件 → useExcelParser 逐行解析 → useExcelValidator 全量校验
    → ExcelPreview 展示（分页 + 错误高亮 + hover提示）
    → 用户通过 ExcelCellEdit 内联修正
    → useExcelSubmitter 分批提交 → 成功回调 / 失败+回滚
```

### 组件职责

**ExcelImport.vue** — 主组件，负责：
- 接收 Props，向下分发
- 编排各 composable 的调用顺序
- 管理全局状态（loading、error、success）
- 透传 Events 给使用者

**ExcelUploader.vue** — 上传区，负责：
- 文件拖拽上传和点击选择
- 文件类型和大小检测
- 模板下载按钮
- 显示已选文件信息

**ExcelPreview.vue** — 数据预览表，负责：
- Element Plus el-table 分页展示
- 错误行整行高亮、错误单元格标红
- 错误详情 tooltip/hover 提示
- 必填列、字段说明展示

**ExcelCellEdit.vue** — 单元格编辑，负责：
- 双击进入编辑模式
- 编辑后触发单格/单行重新校验
- 支持按 ESC 取消编辑

**SubmitBar.vue** — 底部操作栏，负责：
- 总行数/通过数/错误数统计
- 提交按钮（校验未通过时禁用）
- 提交进度条
- 提交结果展示

### Composables

**useExcelParser(file, columns)**
- 使用 ExcelJS 创建 Workbook 对象
- 遍历 worksheet 的 rows，跳过完全空行
- 首行作为表头，按 label 匹配 column.field
- 按 column.type 转换数据类型（异常时收集到 parseErrors 不中断）
- 返回 `ParseResult`

**useExcelValidator(rows, columns)**
- 逐行逐列执行配置的 validators
- required 校验空值
- unique 校验跨行唯一性
- range 校验数字范围
- pattern 校验正则
- enum 校验值是否在列表中
- custom 调用用户自定义函数（支持异步）
- 所有错误收集完成后一次性返回（非 fail-fast）

**useExcelSubmitter(validRows, config)**
- 按 batchSize 将数据切分为多批
- 顺序逐批提交（非并行，避免后端压力）
- 每批提交后更新进度
- 某批失败：调用 rollbackApi 回滚已提交批次
- 返回 `SubmitResult`

**useExcelTemplate(columns, config?)**
- 使用 ExcelJS 新建 Workbook
- 写入说明行（可选）
- 写入表头行（必填列标 `*`）
- 设置列格式（数字/日期/文本）
- enum 列添加数据验证
- 冻结表头
- 生成 Blob，创建下载链接，触发下载

## 7. 核心类型定义

```typescript
interface ColumnConfig {
  label: string
  field: string
  type?: 'string' | 'number' | 'boolean' | 'date'
  required?: boolean
  default?: any
  description?: string
  validators?: ValidatorConfig[]
}

interface ValidatorConfig {
  type: 'required' | 'unique' | 'range' | 'pattern' | 'enum' | 'custom'
  min?: number
  max?: number
  pattern?: string | RegExp
  enum?: any[]
  message?: string
  validate?: (value: any, row: Record<string, any>, allData: Record<string, any>[]) => boolean | string | Promise<boolean | string>
}

interface CellError {
  row: number
  field: string
  value?: any
  message: string
  type: string
}

interface ParseResult {
  headers: string[]
  rows: Record<string, any>[]
  totalRows: number
  parseErrors: CellError[]
}

interface ValidationResult {
  valid: boolean
  total: number
  validCount: number
  errorCount: number
  errors: CellError[]
  rows: RowWithErrors[]
}

interface RowWithErrors {
  data: Record<string, any>
  errors: CellError[]
  valid: boolean
}

interface SubmitConfig {
  submitApi: string | ((data: Record<string, any>[]) => Promise<any>)
  rollbackApi?: string | ((batchId: string) => Promise<any>)
  batchSize?: number
}

interface SubmitProgress {
  current: number
  total: number
  batch: number
  totalBatches: number
}

interface SubmitResult {
  success: boolean
  totalBatches: number
  completedBatches: number
  failedBatch?: number
  rollbackSuccess?: boolean
  error?: string
}
```

## 8. 包构建配置

### package.json

```
- name: @scope/vue-excel-import
- type: module
- vue 和 element-plus 为 peerDependencies
- exceljs 为 dependencies
- Vite lib 模式构建，输出 ESM + UMD + TS 类型
```

### vite.config.ts

- `@vitejs/plugin-vue` 插件
- `build.lib` 模式，入口 `src/index.ts`
- `rollupOptions.external`: vue, element-plus, exceljs
- 类型生成通过 `vue-tsc --noEmit && vite build` 确保类型检查

## 9. 错误处理策略

| 阶段 | 错误类型 | 处理方式 |
|------|----------|----------|
| 文件上传 | 文件过大/格式不支持 | 上传前检测，立即提示 |
| 解析 | 列匹配失败/类型转换失败 | 收集到 parseErrors，不影响整体解析 |
| 校验 | 数据不符合校验规则 | 全量收集，一次性展示所有错误 |
| 提交 | 接口返回错误 | 触发回滚，@error 事件通知 |
| 网络 | 请求超时/断网 | 进度条暂停，提示用户重试 |

## 10. 性能策略

- **流式解析**：使用 ExcelJS 逐行读取 worksheet，不一次性加载整个文件到内存
- **跳过空行**：迭代时检查行内容，完全空行直接跳过
- **分页预览**：预览表格分页展示（默认每页 100 行），避免渲染大量 DOM
- **分批提交**：默认 500 条/批，防止单次请求体过大
- **全部修正后才可提交**：有校验错误的行必须全部修正后才允许提交，提交按钮在存在错误时禁用
- **仅提交有效数据**：提交时只推送校验通过的行，错误行在提交前必须修正或清除
