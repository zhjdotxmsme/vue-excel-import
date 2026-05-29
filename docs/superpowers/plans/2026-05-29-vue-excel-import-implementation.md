# Vue Excel Import Component — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reusable Vue 3 + TypeScript + Element Plus Excel import component package that parses, validates, and batch-submits Excel data from the frontend.

**Architecture:** 4 independent composables (parse/validate/submit/template) handle all logic; 5 Vue components provide the UI layer; the main component `ExcelImport.vue` orchestrates them. Pure-function design makes composables independently testable.

**Tech Stack:** Vue 3.4, TypeScript 5, Element Plus 2.8, ExcelJS 4.4, Vite 5, Vitest 2

---

## File Structure

```
vue-excel-import/
├── package.json
├── tsconfig.json / tsconfig.node.json
├── vite.config.ts / vitest.config.ts
├── env.d.ts / .gitignore
├── src/
│   ├── index.ts
│   ├── types/index.ts
│   ├── utils/excel.ts
│   ├── composables/
│   │   ├── useExcelParser.ts       # + .test.ts
│   │   ├── useExcelValidator.ts    # + .test.ts
│   │   ├── useExcelSubmitter.ts    # + .test.ts
│   │   └── useExcelTemplate.ts     # + .test.ts
│   └── components/
│       ├── ExcelImport.vue
│       ├── ExcelUploader.vue
│       ├── ExcelPreview.vue
│       ├── ExcelCellEdit.vue
│       └── SubmitBar.vue
```

---

### Task 1: Project Scaffolding

**Files:**
- Create: `package.json`, `tsconfig.json`, `tsconfig.node.json`, `vite.config.ts`, `vitest.config.ts`, `env.d.ts`, `.gitignore`

- [ ] **Step 1: Create package.json**

```json
{
  "name": "@scope/vue-excel-import",
  "version": "0.1.0",
  "private": false,
  "type": "module",
  "main": "./dist/vue-excel-import.umd.cjs",
  "module": "./dist/vue-excel-import.js",
  "types": "./dist/index.d.ts",
  "exports": {
    ".": {
      "import": "./dist/vue-excel-import.js",
      "require": "./dist/vue-excel-import.umd.cjs",
      "types": "./dist/index.d.ts"
    },
    "./dist/style.css": "./dist/style.css"
  },
  "files": ["dist"],
  "scripts": {
    "dev": "vite",
    "build": "vue-tsc --noEmit && vite build",
    "test": "vitest run",
    "test:watch": "vitest"
  },
  "peerDependencies": {
    "vue": "^3.4.0",
    "element-plus": "^2.8.0"
  },
  "dependencies": {
    "exceljs": "^4.4.0"
  },
  "devDependencies": {
    "@vitejs/plugin-vue": "^5.0.0",
    "typescript": "~5.4.0",
    "vite": "^5.4.0",
    "vue-tsc": "^2.0.0",
    "vitest": "^2.0.0",
    "jsdom": "^24.0.0"
  }
}
```

- [ ] **Step 2: Create tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "module": "ESNext",
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "preserve",
    "strict": true,
    "noUnusedLocals": false,
    "noUnusedParameters": false,
    "declaration": true,
    "declarationDir": "./dist",
    "paths": { "@/*": ["./src/*"] }
  },
  "include": ["src/**/*.ts", "src/**/*.d.ts", "src/**/*.vue", "env.d.ts"]
}
```

- [ ] **Step 3: Create tsconfig.node.json**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "strict": true
  },
  "include": ["vite.config.ts", "vitest.config.ts"]
}
```

- [ ] **Step 4: Create vite.config.ts**

```typescript
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'

export default defineConfig({
  plugins: [vue()],
  resolve: { alias: { '@': resolve(__dirname, 'src') } },
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'VueExcelImport',
      fileName: 'vue-excel-import'
    },
    rollupOptions: {
      external: ['vue', 'element-plus', 'exceljs'],
      output: {
        globals: {
          vue: 'Vue',
          'element-plus': 'ElementPlus',
          exceljs: 'ExcelJS'
        }
      }
    }
  }
})
```

- [ ] **Step 5: Create vitest.config.ts**

```typescript
import { defineConfig } from 'vitest/config'
import vue from '@vitejs/plugin-vue'

export default defineConfig({
  plugins: [vue()],
  test: {
    environment: 'jsdom',
    globals: true,
    include: ['src/**/*.test.ts']
  }
})
```

- [ ] **Step 6: Create env.d.ts**

```typescript
/// <reference types="vite/client" />
declare module '*.vue' {
  import type { DefineComponent } from 'vue'
  const component: DefineComponent<{}, {}, any>
  export default component
}
```

- [ ] **Step 7: Create .gitignore**

```
node_modules/
dist/
*.local
.superpowers/
```

- [ ] **Step 8: Install dependencies**

Run: `npm install`
Expected: All dependencies installed without errors.

- [ ] **Step 9: Create src directory structure**

Run: `mkdir -p src/types src/utils src/composables src/components`

- [ ] **Step 10: Commit**

```bash
git add -A && git commit -m "chore: scaffold project with Vite + Vue 3 + TypeScript"
```

---

### Task 2: Core Types

**Files:**
- Create: `src/types/index.ts`

- [ ] **Step 1: Create type definitions**

```typescript
export type ColumnType = 'string' | 'number' | 'boolean' | 'date'

export interface ColumnConfig {
  label: string
  field: string
  type?: ColumnType
  required?: boolean
  default?: any
  description?: string
  validators?: ValidatorConfig[]
}

export interface ValidatorConfig {
  type: 'required' | 'unique' | 'range' | 'pattern' | 'enum' | 'custom'
  min?: number
  max?: number
  pattern?: string | RegExp
  enum?: any[]
  message?: string
  validate?: (
    value: any,
    row: Record<string, any>,
    allData: Record<string, any>[]
  ) => boolean | string | Promise<boolean | string>
}

export interface CellError {
  row: number
  field: string
  value?: any
  message: string
  type: string
}

export interface ParseResult {
  headers: string[]
  rows: Record<string, any>[]
  totalRows: number
  parseErrors: CellError[]
}

export interface RowWithErrors {
  data: Record<string, any>
  errors: CellError[]
  valid: boolean
}

export interface ValidationResult {
  valid: boolean
  total: number
  validCount: number
  errorCount: number
  errors: CellError[]
  rows: RowWithErrors[]
}

export interface SubmitConfig {
  submitApi: string | ((data: Record<string, any>[]) => Promise<any>)
  rollbackApi?: string | ((batchId: string) => Promise<any>)
  batchSize?: number
}

export interface SubmitProgress {
  current: number
  total: number
  batch: number
  totalBatches: number
}

export interface SubmitResult {
  success: boolean
  totalBatches: number
  completedBatches: number
  failedBatch?: number
  rollbackSuccess?: boolean
  error?: string
}
```

- [ ] **Step 2: Verify LSP clean**

Run: `npx vue-tsc --noEmit src/types/index.ts` (or just check no red squigglies)
Expected: No errors

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add core type definitions"
```

---

### Task 3: Utility Functions (excel.ts)

**Files:**
- Create: `src/utils/excel.ts`
- Create: `src/utils/excel.test.ts`

- [ ] **Step 1: Write failing tests**

```typescript
import { describe, it, expect } from 'vitest'
import { excelSerialDateToJSDate, formatDate, convertValue } from './excel'

describe('excelSerialDateToJSDate', () => {
  it('converts Excel serial 45291 to 2024-01-01', () => {
    const d = excelSerialDateToJSDate(45291)
    expect(d.getFullYear()).toBe(2024)
    expect(d.getMonth()).toBe(0)
    expect(d.getDate()).toBe(1)
  })
})

describe('formatDate', () => {
  it('formats Date to YYYY-MM-DD', () => {
    expect(formatDate(new Date(2024, 0, 15))).toBe('2024-01-15')
  })
})

describe('convertValue', () => {
  describe('type=string', () => {
    it('trims whitespace', () => expect(convertValue('  hi  ', 'string').value).toBe('hi'))
    it('converts number to string', () => expect(convertValue(25, 'string').value).toBe('25'))
    it('converts boolean', () => expect(convertValue(true, 'string').value).toBe('true'))
    it('converts Date to YYYY-MM-DD', () => expect(convertValue(new Date(2024, 0, 15), 'string').value).toBe('2024-01-15'))
    it('returns empty string for null', () => expect(convertValue(null, 'string').value).toBe(''))
  })

  describe('type=number', () => {
    it('keeps number', () => expect(convertValue(42, 'number').value).toBe(42))
    it('parses numeric string', () => expect(convertValue('42', 'number').value).toBe(42))
    it('converts boolean to 1/0', () => { expect(convertValue(true, 'number').value).toBe(1); expect(convertValue(false, 'number').value).toBe(0) })
    it('errors on non-numeric string', () => expect(convertValue('abc', 'number').error).toBeTruthy())
  })

  describe('type=boolean', () => {
    it('keeps boolean', () => { expect(convertValue(true, 'boolean').value).toBe(true) })
    it('0=false, non-0=true', () => { expect(convertValue(0, 'boolean').value).toBe(false); expect(convertValue(1, 'boolean').value).toBe(true) })
    it('"true"/"1"/"yes"=true', () => { expect(convertValue('true', 'boolean').value).toBe(true); expect(convertValue('yes', 'boolean').value).toBe(true) })
    it('errors on ambiguous string', () => expect(convertValue('maybe', 'boolean').error).toBeTruthy())
  })

  describe('type=date', () => {
    it('formats Date object', () => expect(convertValue(new Date(2024, 5, 15), 'date').value).toBe('2024-06-15'))
    it('Unix timestamp (number > 100000)', () => expect(convertValue(1704067200000, 'date').value).toBe('2024-01-01'))
    it('Excel serial (number <= 100000)', () => expect(convertValue(45291, 'date').value).toBe('2024-01-01'))
    it('parses date string', () => expect(convertValue('2024-06-15', 'date').value).toBe('2024-06-15'))
    it('errors on invalid string', () => expect(convertValue('not-date', 'date').error).toBeTruthy())
  })
})
```

- [ ] **Step 2: Run to verify failures**

Run: `npx vitest run src/utils/excel.test.ts`
Expected: Fails with import errors.

- [ ] **Step 3: Write implementation**

```typescript
import type { ColumnType } from '../types'

const EXCEL_EPOCH_SECONDS = 25569

export function excelSerialDateToJSDate(serial: number): Date {
  return new Date((serial - EXCEL_EPOCH_SECONDS) * 86400000)
}

export function formatDate(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

function extractCellValue(raw: any): any {
  if (raw === null || raw === undefined) return null
  if (typeof raw !== 'object') return raw
  if ('richText' in raw) return raw.richText.map((r: any) => r.text).join('')
  if ('error' in raw) return raw
  if ('hyperlink' in raw) return raw.text ?? String(raw.hyperlink)
  return raw
}

export function convertValue(raw: any, type: ColumnType): { value: any; error?: string } {
  const v = extractCellValue(raw)
  if (v === null || v === undefined) return type === 'string' ? { value: '' } : { value: null }

  // Excel error object
  if (typeof v === 'object' && v !== null && 'error' in v) {
    if (type === 'string') return { value: String(v.error) }
    return { value: null, error: `Excel error: ${v.error}` }
  }

  switch (type) {
    case 'string': {
      if (typeof v === 'string') return { value: v.trim() }
      if (typeof v === 'number') return { value: Number.isInteger(v) ? String(v) : String(v) }
      if (typeof v === 'boolean') return { value: String(v) }
      if (v instanceof Date) return { value: formatDate(v) }
      return { value: String(v) }
    }

    case 'number': {
      if (typeof v === 'number') return { value: v }
      if (typeof v === 'boolean') return { value: v ? 1 : 0 }
      if (typeof v === 'string') {
        const t = v.trim()
        if (!t) return { value: null }
        const n = Number(t)
        return isNaN(n) ? { value: v, error: `"${v}" 无法转换为数字` } : { value: n }
      }
      if (v instanceof Date) return { value: v.getTime() }
      return { value: v, error: `无法将 ${typeof v} 转换为数字` }
    }

    case 'boolean': {
      if (typeof v === 'boolean') return { value: v }
      if (typeof v === 'number') return { value: v !== 0 }
      if (typeof v === 'string') {
        const l = v.trim().toLowerCase()
        if (['true', '1', 'yes'].includes(l)) return { value: true }
        if (['false', '0', 'no'].includes(l)) return { value: false }
        return { value: v, error: `"${v}" 无法转换为布尔值` }
      }
      return { value: v, error: `无法将 ${typeof v} 转换为布尔值` }
    }

    case 'date': {
      if (v instanceof Date) {
        return isNaN(v.getTime()) ? { value: v, error: '无效日期' } : { value: formatDate(v) }
      }
      if (typeof v === 'number') {
        const d = v > 100000 ? new Date(v) : excelSerialDateToJSDate(v)
        return isNaN(d.getTime()) ? { value: v, error: `无效日期数值: ${v}` } : { value: formatDate(d) }
      }
      if (typeof v === 'string') {
        if (!v.trim()) return { value: null }
        const d = new Date(v.trim())
        return isNaN(d.getTime()) ? { value: v, error: `"${v}" 无法解析为日期` } : { value: formatDate(d) }
      }
      return { value: v, error: `无法将 ${typeof v} 转换为日期` }
    }
  }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/utils/excel.test.ts`
Expected: All tests pass.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: add Excel utility functions (serial date conversion, type conversion)"
```

---

### Task 4: useExcelParser Composable

**Files:**
- Create: `src/composables/useExcelParser.ts`
- Create: `src/composables/useExcelParser.test.ts`

- [ ] **Step 1: Write failing tests (with in-memory ExcelJS workbook)**

```typescript
import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import type { ColumnConfig } from '../types'
import { useExcelParser } from './useExcelParser'

async function makeFile(headers: string[], rows: any[][]): Promise<File> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Sheet1')
  ws.addRow(headers)
  rows.forEach(r => ws.addRow(r))
  const buf = await wb.xlsx.writeBuffer()
  return new File([buf], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

describe('useExcelParser', () => {
  const columns: ColumnConfig[] = [
    { label: '姓名', field: 'userName', type: 'string' },
    { label: '年龄', field: 'age', type: 'number' },
    { label: '入职日期', field: 'hireDate', type: 'date' }
  ]

  it('parses valid Excel and maps columns by label', async () => {
    const file = await makeFile(
      ['姓名', '年龄', '入职日期'],
      [['张三', 28, new Date(2024, 0, 15)], ['李四', 35, new Date(2023, 5, 1)]]
    )
    const result = await useExcelParser(file, columns)
    expect(result.totalRows).toBe(2)
    expect(result.rows[0].userName).toBe('张三')
    expect(result.rows[0].age).toBe(28)
    expect(result.rows[0].hireDate).toBe('2024-01-15')
    expect(result.parseErrors).toHaveLength(0)
  })

  it('skips completely empty rows', async () => {
    const file = await makeFile(
      ['姓名', '年龄'],
      [['张三', 28], [], ['李四', 35]]
    )
    const result = await useExcelParser(file, columns)
    expect(result.totalRows).toBe(2)
  })

  it('ignores Excel columns not in config', async () => {
    const file = await makeFile(
      ['姓名', '年龄', '备注'],
      [['张三', 28, 'abc']]
    )
    const result = await useExcelParser(file, columns)
    expect(result.totalRows).toBe(1)
    expect(Object.keys(result.rows[0])).not.toContain('备注')
  })

  it('collects parse errors for invalid type conversion', async () => {
    const file = await makeFile(
      ['姓名', '年龄'],
      [['张三', 'not-a-number']]
    )
    const result = await useExcelParser(file, columns)
    expect(result.parseErrors.length).toBeGreaterThan(0)
    expect(result.parseErrors[0].type).toBe('type-conversion')
  })

  it('returns empty result when no columns match Excel headers', async () => {
    const file = await makeFile(['A', 'B'], [['x', 'y']])
    const result = await useExcelParser(file, columns)
    expect(result.totalRows).toBe(1)
    expect(Object.keys(result.rows[0])).toHaveLength(0)
  })

  it('handles empty Excel file (header only)', async () => {
    const file = await makeFile(['姓名', '年龄'], [])
    const result = await useExcelParser(file, columns)
    expect(result.totalRows).toBe(0)
  })
})
```

- [ ] **Step 2: Run to verify failures**

Run: `npx vitest run src/composables/useExcelParser.test.ts`
Expected: Fails — module not found.

- [ ] **Step 3: Write implementation**

```typescript
import ExcelJS from 'exceljs'
import type { ColumnConfig, ParseResult, CellError } from '../types'
import { convertValue } from '../utils/excel'

export async function useExcelParser(
  file: File,
  columns: ColumnConfig[]
): Promise<ParseResult> {
  const workbook = new ExcelJS.Workbook()
  const buffer = await file.arrayBuffer()
  await workbook.xlsx.load(buffer)

  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    return {
      headers: [],
      rows: [],
      totalRows: 0,
      parseErrors: [{ row: 0, field: '', value: null, message: '工作表不存在', type: 'parse' }]
    }
  }

  const errors: CellError[] = []
  const rows: Record<string, any>[] = []
  let headers: string[] = []
  let headerIndexMap = new Map<number, ColumnConfig>()

  // Read header row (row 1) and build index → config mapping
  const headerRow = worksheet.getRow(1)
  headerRow.eachCell((cell, colIdx) => {
    const label = String(cell.value ?? '').trim()
    headers.push(label)
    const config = columns.find(c => c.label === label)
    if (config) headerIndexMap.set(colIdx, config)
  })

  // Iterate data rows (starting from row 2)
  worksheet.eachRow((row, rowIdx) => {
    if (rowIdx === 1) return

    let hasContent = false
    const rowData: Record<string, any> = {}

    row.eachCell((cell, colIdx) => {
      if (cell.value !== null && cell.value !== undefined && cell.value !== '') {
        hasContent = true
      }
      const config = headerIndexMap.get(colIdx)
      if (!config) return

      const result = convertValue(cell.value, config.type ?? 'string')
      rowData[config.field] = result.value

      if (result.error) {
        errors.push({
          row: rowIdx,
          field: config.field,
          value: cell.value,
          message: result.error,
          type: 'type-conversion'
        })
      }
    })

    if (hasContent) rows.push(rowData)
  })

  return { headers, rows, totalRows: rows.length, parseErrors: errors }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/composables/useExcelParser.test.ts`
Expected: All tests pass.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: add useExcelParser composable with ExcelJS streaming"
```

---

### Task 5: useExcelValidator Composable

**Files:**
- Create: `src/composables/useExcelValidator.ts`
- Create: `src/composables/useExcelValidator.test.ts`

- [ ] **Step 1: Write failing tests**

```typescript
import { describe, it, expect } from 'vitest'
import type { ColumnConfig } from '../types'
import { useExcelValidator } from './useExcelValidator'

describe('useExcelValidator', () => {
  it('returns valid=true when no validators configured', () => {
    const columns: ColumnConfig[] = [
      { label: 'Name', field: 'name', type: 'string' }
    ]
    const result = useExcelValidator([{ name: 'Alice' }], columns)
    expect(result.valid).toBe(true)
    expect(result.validCount).toBe(1)
    expect(result.errorCount).toBe(0)
  })

  it('validates required fields', () => {
    const columns: ColumnConfig[] = [
      { label: 'Name', field: 'name', type: 'string', required: true }
    ]
    const result = useExcelValidator([{ name: '' }, { name: 'Alice' }], columns)
    expect(result.valid).toBe(false)
    expect(result.validCount).toBe(1)
    expect(result.errorCount).toBe(1)
    expect(result.errors[0].type).toBe('required')
  })

  it('validates unique fields', () => {
    const columns: ColumnConfig[] = [
      { label: 'Email', field: 'email', type: 'string', validators: [{ type: 'unique' }] }
    ]
    const result = useExcelValidator(
      [{ email: 'a@b.com' }, { email: 'a@b.com' }, { email: 'c@d.com' }],
      columns
    )
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(2) // both duplicates have errors
  })

  it('validates range (number)', () => {
    const columns: ColumnConfig[] = [
      { label: 'Age', field: 'age', type: 'number', validators: [{ type: 'range', min: 0, max: 150 }] }
    ]
    const result = useExcelValidator([{ age: 25 }, { age: -1 }, { age: 200 }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(2)
  })

  it('validates pattern (regex)', () => {
    const columns: ColumnConfig[] = [
      { label: 'Email', field: 'email', type: 'string', validators: [{ type: 'pattern', pattern: '^\\S+@\\S+\\.\\S+$' }] }
    ]
    const result = useExcelValidator([{ email: 'good@test.com' }, { email: 'bad' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
  })

  it('validates enum values', () => {
    const columns: ColumnConfig[] = [
      { label: 'Status', field: 'status', type: 'string', validators: [{ type: 'enum', enum: ['active', 'inactive'] }] }
    ]
    const result = useExcelValidator([{ status: 'active' }, { status: 'unknown' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
  })

  it('validates with custom sync function', () => {
    const columns: ColumnConfig[] = [
      {
        label: 'Code', field: 'code', type: 'string',
        validators: [{ type: 'custom', validate: (v) => v?.startsWith('CODE-') ? true : '必须以 CODE- 开头' }]
      }
    ]
    const result = useExcelValidator([{ code: 'CODE-123' }, { code: 'BAD' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
    expect(result.errors[0].message).toBe('必须以 CODE- 开头')
  })

  it('validates with custom async function', async () => {
    const columns: ColumnConfig[] = [
      {
        label: 'Email', field: 'email', type: 'string',
        validators: [{
          type: 'custom',
          validate: async (v) => {
            await new Promise(r => setTimeout(r, 10))
            return v?.includes('@') ? true : 'invalid'
          }
        }]
      }
    ]
    const result = await useExcelValidator([{ email: 'a@b.com' }, { email: 'abc' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
  })

  it('handles null/empty values in validators', () => {
    const columns: ColumnConfig[] = [
      { label: 'Age', field: 'age', type: 'number', validators: [{ type: 'range', min: 0, max: 150 }] }
    ]
    const result = useExcelValidator([{ age: null }, { age: '' }], columns)
    expect(result.valid).toBe(true) // null/empty skips non-required validators
  })
})
```

- [ ] **Step 2: Run to verify failures**

Run: `npx vitest run src/composables/useExcelValidator.test.ts`
Expected: Fails.

- [ ] **Step 3: Write implementation**

```typescript
import type { ColumnConfig, ValidationResult, RowWithErrors, CellError, ValidatorConfig } from '../types'

function runValidator(
  value: any,
  row: Record<string, any>,
  allData: Record<string, any>[],
  config: ValidatorConfig,
  field: string,
  rowIdx: number
): CellError | null {
  const message = config.message ?? ''

  switch (config.type) {
    case 'required': {
      if (value === null || value === undefined || value === '') {
        return { row: rowIdx, field, value, message: message || '此字段为必填项', type: 'required' }
      }
      return null
    }

    case 'unique': {
      const duplicates = allData.filter(d => d[field] === value && d !== row)
      if (duplicates.length > 0) {
        return { row: rowIdx, field, value, message: message || `值 "${value}" 已存在`, type: 'unique' }
      }
      return null
    }

    case 'range': {
      if (value === null || value === undefined || value === '') return null
      const num = Number(value)
      if (isNaN(num)) return { row: rowIdx, field, value, message: message || '不是有效数字', type: 'range' }
      if (config.min !== undefined && num < config.min) {
        return { row: rowIdx, field, value, message: message || `不能小于 ${config.min}`, type: 'range' }
      }
      if (config.max !== undefined && num > config.max) {
        return { row: rowIdx, field, value, message: message || `不能大于 ${config.max}`, type: 'range' }
      }
      return null
    }

    case 'pattern': {
      if (value === null || value === undefined || value === '') return null
      const pattern = config.pattern instanceof RegExp ? config.pattern : new RegExp(config.pattern!)
      if (!pattern.test(String(value))) {
        return { row: rowIdx, field, value, message: message || '格式不匹配', type: 'pattern' }
      }
      return null
    }

    case 'enum': {
      if (value === null || value === undefined || value === '') return null
      if (!config.enum!.includes(value)) {
        return { row: rowIdx, field, value, message: message || `值 "${value}" 不在允许列表中`, type: 'enum' }
      }
      return null
    }

    case 'custom': {
      if (value === null || value === undefined || value === '') return null
      const result = config.validate!(value, row, allData)
      if (result === true || result === '') return null
      const msg = typeof result === 'string' ? result : (message || '自定义校验未通过')
      return { row: rowIdx, field, value, message: msg, type: 'custom' }
    }

    default:
      return null
  }
}

export function useExcelValidator(
  rows: Record<string, any>[],
  columns: ColumnConfig[]
): ValidationResult {
  const errors: CellError[] = []

  const rowWrappers: RowWithErrors[] = rows.map((data, idx) => {
    const rowErrors: CellError[] = []
    const rowNum = idx + 1 // 1-indexed display row

    for (const col of columns) {
      const value = data[col.field]

      // Required check (from ColumnConfig.required)
      if (col.required && (value === null || value === undefined || value === '')) {
        rowErrors.push({
          row: rowNum,
          field: col.field,
          value,
          message: '此字段为必填项',
          type: 'required'
        })
        // Don't run other validators on missing required fields
        continue
      }

      // Run configured validators
      if (col.validators) {
        for (const vc of col.validators) {
          // No need to re-check required here; already done above
          if (vc.type === 'required') continue

          const err = runValidator(value, data, rows, vc, col.field, rowNum)
          if (err) rowErrors.push(err)
        }
      }
    }

    errors.push(...rowErrors)
    return { data, errors: rowErrors, valid: rowErrors.length === 0 }
  })

  const validCount = rowWrappers.filter(r => r.valid).length

  return {
    valid: validCount === rows.length,
    total: rows.length,
    validCount,
    errorCount: rows.length - validCount,
    errors,
    rows: rowWrappers
  }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/composables/useExcelValidator.test.ts`
Expected: All tests pass.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: add useExcelValidator composable with 6 validator types"
```

---

### Task 6: useExcelSubmitter Composable

**Files:**
- Create: `src/composables/useExcelSubmitter.ts`
- Create: `src/composables/useExcelSubmitter.test.ts`

- [ ] **Step 1: Write failing tests**

```typescript
import { describe, it, expect, vi } from 'vitest'
import { useExcelSubmitter } from './useExcelSubmitter'

describe('useExcelSubmitter', () => {
  it('submits all rows in one batch when under batch size', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const result = await useExcelSubmitter(
      [{ name: 'A' }, { name: 'B' }],
      { submitApi: fn, batchSize: 500 }
    )
    expect(fn).toHaveBeenCalledTimes(1)
    expect(fn).toHaveBeenCalledWith([{ name: 'A' }, { name: 'B' }])
    expect(result.success).toBe(true)
    expect(result.totalBatches).toBe(1)
  })

  it('splits into multiple batches when over batch size', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const data = Array.from({ length: 1200 }, (_, i) => ({ id: i + 1 }))
    const result = await useExcelSubmitter(data, { submitApi: fn, batchSize: 500 })
    expect(fn).toHaveBeenCalledTimes(3)
    expect(fn.mock.calls[0][0]).toHaveLength(500)
    expect(fn.mock.calls[1][0]).toHaveLength(500)
    expect(fn.mock.calls[2][0]).toHaveLength(200)
    expect(result.success).toBe(true)
  })

  it('calls rollbackApi on failure and returns rollbackSuccess', async () => {
    const submitFn = vi.fn()
      .mockResolvedValueOnce({ ok: true })
      .mockRejectedValueOnce(new Error('batch failed'))
    const rollbackFn = vi.fn().mockResolvedValue({ ok: true })
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i + 1 }))

    const result = await useExcelSubmitter(data, {
      submitApi: submitFn,
      rollbackApi: rollbackFn,
      batchSize: 500
    })

    expect(submitFn).toHaveBeenCalledTimes(2)
    expect(rollbackFn).toHaveBeenCalledTimes(1)
    expect(result.success).toBe(false)
    expect(result.failedBatch).toBe(2)
    expect(result.rollbackSuccess).toBe(true)
  })

  it('reports progress via callback', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const progresses: any[] = []
    const onProgress = (p: any) => progresses.push(p)
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i + 1 }))

    await useExcelSubmitter(data, {
      submitApi: fn,
      batchSize: 500,
      onProgress
    })

    expect(progresses).toHaveLength(2)
    expect(progresses[0]).toMatchObject({ batch: 1, totalBatches: 2 })
    expect(progresses[1]).toMatchObject({ batch: 2, totalBatches: 2 })
  })

  it('calls submitApi as string URL via fetch', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({ ok: true, json: () => Promise.resolve({}) })
    const data = [{ name: 'A' }]
    const result = await useExcelSubmitter(data, {
      submitApi: '/api/import',
      batchSize: 500
    })
    expect(globalThis.fetch).toHaveBeenCalledWith('/api/import', expect.objectContaining({
      method: 'POST',
      body: expect.any(String)
    }))
    expect(result.success).toBe(true)
  })
})
```

- [ ] **Step 2: Run to verify failures**

Run: `npx vitest run src/composables/useExcelSubmitter.test.ts`
Expected: Fails.

- [ ] **Step 3: Write implementation**

```typescript
import type { SubmitConfig, SubmitResult, SubmitProgress } from '../types'

interface ExtendedSubmitConfig extends SubmitConfig {
  onProgress?: (progress: SubmitProgress) => void
}

export async function useExcelSubmitter(
  rows: Record<string, any>[],
  config: ExtendedSubmitConfig
): Promise<SubmitResult> {
  const { submitApi, rollbackApi } = config
  const batchSize = config.batchSize ?? 500
  const totalBatches = Math.ceil(rows.length / batchSize)
  const completedBatches: any[] = []

  async function sendBatch(batch: Record<string, any>[]): Promise<any> {
    if (typeof submitApi === 'function') {
      return submitApi(batch)
    }
    const response = await fetch(submitApi, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(batch)
    })
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return response.json()
  }

  async function rollback(): Promise<boolean> {
    if (!rollbackApi) return false
    try {
      if (typeof rollbackApi === 'function') {
        await rollbackApi(String(Date.now()))
      } else {
        await fetch(rollbackApi, { method: 'POST' })
      }
      return true
    } catch {
      return false
    }
  }

  for (let i = 0; i < totalBatches; i++) {
    const start = i * batchSize
    const batch = rows.slice(start, start + batchSize)
    const batchNum = i + 1

    try {
      const result = await sendBatch(batch)
      completedBatches.push(result)

      config.onProgress?.({
        current: Math.min((i + 1) * batchSize, rows.length),
        total: rows.length,
        batch: batchNum,
        totalBatches
      })
    } catch (err: any) {
      const rollbackSuccess = await rollback()
      return {
        success: false,
        totalBatches,
        completedBatches: i,
        failedBatch: batchNum,
        rollbackSuccess,
        error: err.message || '提交失败'
      }
    }
  }

  return {
    success: true,
    totalBatches,
    completedBatches: totalBatches,
    rollbackSuccess: true
  }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/composables/useExcelSubmitter.test.ts`
Expected: All tests pass.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: add useExcelSubmitter composable with batch and rollback"
```

---

### Task 7: useExcelTemplate Composable

**Files:**
- Create: `src/composables/useExcelTemplate.ts`
- Create: `src/composables/useExcelTemplate.test.ts`

- [ ] **Step 1: Write failing tests**

```typescript
import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import type { ColumnConfig } from '../types'
import { useExcelTemplate } from './useExcelTemplate'

describe('useExcelTemplate', () => {
  const columns: ColumnConfig[] = [
    { label: '姓名', field: 'userName', type: 'string', required: true },
    { label: '年龄', field: 'age', type: 'number' },
    { label: '部门', field: 'dept', type: 'string', validators: [{ type: 'enum', enum: ['技术部', '市场部', '财务部'] }] },
    { label: '入职日期', field: 'hireDate', type: 'date' }
  ]

  it('generates a downloadable Blob', async () => {
    const blob = await useExcelTemplate(columns, { fileName: 'test.xlsx' })
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  })

  it('creates workbook with correct headers', async () => {
    const blob = await useExcelTemplate(columns)
    const workbook = new ExcelJS.Workbook()
    const buf = await blob.arrayBuffer()
    await workbook.xlsx.load(buf)
    const ws = workbook.worksheets[0]
    expect(ws).toBeDefined()
    const row1 = ws.getRow(1)
    expect(row1.getCell(1).value).toContain('姓名')
    expect(row1.getCell(2).value).toContain('年龄')
  })

  it('marks required columns with asterisk in header', async () => {
    const blob = await useExcelTemplate(columns)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blob.arrayBuffer())
    const ws = workbook.worksheets[0]
    expect(String(ws.getRow(1).getCell(1).value)).toContain('*')
  })
})
```

- [ ] **Step 2: Run to verify failures**

Run: `npx vitest run src/composables/useExcelTemplate.test.ts`
Expected: Fails.

- [ ] **Step 3: Write implementation**

```typescript
import ExcelJS from 'exceljs'
import type { ColumnConfig } from '../types'

interface TemplateOptions {
  fileName?: string
  sheetName?: string
  description?: string
  exampleRows?: number
}

export async function useExcelTemplate(
  columns: ColumnConfig[],
  options?: TemplateOptions
): Promise<Blob> {
  const workbook = new ExcelJS.Workbook()
  const sheetName = options?.sheetName ?? 'Sheet1'
  const ws = workbook.addWorksheet(sheetName)

  // Optional description row
  if (options?.description) {
    ws.addRow([options.description])
  }

  // Header row
  const headerRow = columns.map(col => {
    const label = col.required ? `${col.label} *` : col.label
    return label
  })
  ws.addRow(headerRow)

  // Style header row
  const headerRowIndex = options?.description ? 2 : 1
  const headerExcelRow = ws.getRow(headerRowIndex)
  headerExcelRow.font = { bold: true }
  headerExcelRow.height = 28

  // Optional example rows (data type hints)
  if (options?.exampleRows && options.exampleRows > 0) {
    for (let i = 0; i < options.exampleRows; i++) {
      const exampleRow = columns.map(col => {
        switch (col.type) {
          case 'number': return 0
          case 'date': return new Date()
          case 'boolean': return true
          default: return ''
        }
      })
      ws.addRow(exampleRow)
    }
  }

  // Set column widths and formats
  columns.forEach((col, idx) => {
    const colIdx = idx + 1
    ws.getColumn(colIdx).width = 16

    switch (col.type) {
      case 'number':
        ws.getColumn(colIdx).numFmt = '0'
        break
      case 'date':
        ws.getColumn(colIdx).numFmt = 'yyyy-mm-dd'
        break
    }

    // Add data validation for enum columns
    const enumValidator = col.validators?.find(v => v.type === 'enum')
    if (enumValidator?.enum && enumValidator.enum.length > 0) {
      const lastRow = ws.rowCount
      ws.dataValidations.add(`${String.fromCharCode(64 + colIdx)}${headerRowIndex + 1}:${String.fromCharCode(64 + colIdx)}${lastRow}`, {
        type: 'list',
        formulae: [enumValidator.enum.join(',')],
        showErrorMessage: true,
        errorTitle: '输入错误',
        error: `请选择: ${enumValidator.enum.join(', ')}`
      })
    }
  })

  // Freeze header row
  ws.views = [{ state: 'frozen', ySplit: headerRowIndex }]

  // Generate blob
  const buffer = await workbook.xlsx.writeBuffer()
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
}

export function downloadTemplateBlob(blob: Blob, fileName: string = '导入模板.xlsx'): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `npx vitest run src/composables/useExcelTemplate.test.ts`
Expected: All tests pass.

- [ ] **Step 5: Commit**

```bash
git add -A && git commit -m "feat: add useExcelTemplate composable with download"
```

---

### Task 8: ExcelUploader.vue Component

**Files:**
- Create: `src/components/ExcelUploader.vue`

- [ ] **Step 1: Create component**

```vue
<template>
  <div class="excel-uploader">
    <div
      class="upload-area"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
      @click="triggerInput"
      :class="{ 'is-dragging': dragging }"
    >
      <div v-if="!file" class="upload-placeholder">
        <div class="upload-icon">📁</div>
        <p class="upload-text">拖拽 Excel 文件到此处，或点击上传</p>
        <p class="upload-hint">支持 .xlsx / .xls</p>
        <el-button type="primary" size="large" @click.stop="triggerInput">
          选择文件
        </el-button>
      </div>

      <div v-else class="file-info">
        <div class="file-icon">{{ fileIcon }}</div>
        <div class="file-details">
          <p class="file-name">{{ file.name }}</p>
          <p class="file-size">{{ formatSize(file.size) }}</p>
        </div>
        <el-button size="small" type="danger" plain @click.stop="clearFile">移除</el-button>
      </div>
    </div>

    <input
      ref="fileInput"
      type="file"
      accept=".xlsx,.xls"
      style="display: none"
      @change="onFileChange"
    />

    <div v-if="showTemplateButton" class="template-download">
      <el-button text @click="emit('download-template')">
        📋 下载导入模板
      </el-button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'

const props = withDefaults(defineProps<{
  showTemplateButton?: boolean
}>(), {
  showTemplateButton: true
})

const emit = defineEmits<{
  (e: 'file-selected', file: File): void
  (e: 'file-cleared'): void
  (e: 'download-template'): void
}>()

const dragging = ref(false)
const file = ref<File | null>(null)
const fileInput = ref<HTMLInputElement | null>(null)

const fileIcon = computed(() => {
  if (!file.value) return ''
  return file.value.name.endsWith('.xls') ? '📄' : '📊'
})

function triggerInput() {
  fileInput.value?.click()
}

function onFileChange(e: Event) {
  const input = e.target as HTMLInputElement
  if (input.files?.length) {
    selectFile(input.files[0])
  }
}

function onDrop(e: DragEvent) {
  dragging.value = false
  const files = e.dataTransfer?.files
  if (files?.length) {
    selectFile(files[0])
  }
}

function selectFile(f: File) {
  if (!f.name.match(/\.xlsx?$/i)) {
    ElMessage.warning('请选择 .xlsx 或 .xls 文件')
    return
  }
  file.value = f
  emit('file-selected', f)
}

function clearFile() {
  file.value = null
  emit('file-cleared')
  if (fileInput.value) fileInput.value.value = ''
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}
</script>

<style scoped>
.excel-uploader {
  margin-bottom: 16px;
}

.upload-area {
  border: 2px dashed #d9d9d9;
  border-radius: 12px;
  padding: 32px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s;
}

.upload-area:hover,
.upload-area.is-dragging {
  border-color: #409eff;
  background: rgba(64, 158, 255, 0.04);
}

.upload-icon { font-size: 40px; margin-bottom: 8px; }
.upload-text { font-weight: 600; margin: 4px 0; }
.upload-hint { font-size: 13px; color: #909399; margin: 2px 0 16px; }

.file-info {
  display: flex;
  align-items: center;
  gap: 12px;
}

.file-icon { font-size: 32px; }
.file-details { flex: 1; text-align: left; }
.file-name { font-weight: 600; margin: 0; }
.file-size { font-size: 12px; color: #909399; margin: 2px 0 0; }

.template-download {
  margin-top: 8px;
  text-align: center;
}
</style>
```

- [ ] **Step 2: Verify LSP diagnostics clean**

Run LSP diagnostics on `src/components/ExcelUploader.vue`
Expected: No errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add ExcelUploader component (drag & drop + template button)"
```

---

### Task 9: ExcelPreview.vue Component

**Files:**
- Create: `src/components/ExcelPreview.vue`

- [ ] **Step 1: Create component**

```vue
<template>
  <div class="excel-preview" v-if="rows.length > 0">
    <div class="preview-header">
      <span class="preview-title">数据预览</span>
      <span class="preview-count">共 {{ total }} 条记录</span>
    </div>

    <el-table
      :data="paginatedRows"
      border
      stripe
      max-height="480"
      :cell-style="cellStyle"
      size="small"
      highlight-current-row
    >
      <el-table-column type="index" label="#" width="50" fixed />
      <el-table-column
        v-for="col in visibleColumns"
        :key="col.field"
        :prop="col.field"
        :label="col.label"
        :width="colWidth(col)"
        show-overflow-tooltip
      >
        <template #default="{ row: r, $index }">
          <div
            class="cell-wrapper"
            :class="{ 'cell-error': hasCellError(r, col.field), 'cell-required': col.required }"
            @dblclick="startEdit(r, col.field, $index)"
          >
            <span>{{ r.data[col.field] ?? '' }}</span>
            <el-tooltip
              v-if="hasCellError(r, col.field)"
              :content="getCellErrors(r, col.field).map(e => e.message).join('; ')"
              placement="top"
              effect="dark"
            >
              <span class="error-badge">!</span>
            </el-tooltip>
          </div>
        </template>
      </el-table-column>

      <el-table-column label="校验状态" width="130" fixed="right">
        <template #default="{ row: r }">
          <el-tag v-if="r.valid" type="success" size="small">通过</el-tag>
          <el-tag v-else type="danger" size="small">
            {{ r.errors.length }} 个错误
          </el-tag>
        </template>
      </el-table-column>
    </el-table>

    <div class="pagination-wrapper" v-if="rows.length > pageSize">
      <el-pagination
        v-model:current-page="currentPage"
        :page-size="pageSize"
        :total="rows.length"
        layout="prev, pager, next"
        small
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import type { ColumnConfig, RowWithErrors } from '../types'

const props = defineProps<{
  columns: ColumnConfig[]
  rows: RowWithErrors[]
  parseErrors: any[]
}>()

const emit = defineEmits<{
  (e: 'edit-cell', rowIndex: number, field: string): void
}>()

const currentPage = ref(1)
const pageSize = 100

const visibleColumns = computed(() => props.columns)

const paginatedRows = computed(() => {
  const start = (currentPage.value - 1) * pageSize
  return props.rows.slice(start, start + pageSize)
})

function colWidth(col: ColumnConfig): number | undefined {
  switch (col.type) {
    case 'number': return 100
    case 'date': return 120
    case 'boolean': return 90
    default: return undefined
  }
}

function hasCellError(row: RowWithErrors, field: string): boolean {
  return row.errors.some(e => e.field === field)
}

function getCellErrors(row: RowWithErrors, field: string) {
  return row.errors.filter(e => e.field === field)
}

function cellStyle({ row, column }: any) {
  const field = column.property
  if (!field) return {}
  const r = row as RowWithErrors
  if (r.errors?.some(e => e.field === field)) {
    return { backgroundColor: '#fef2f2', color: '#dc2626' }
  }
  return {}
}

function startEdit(row: RowWithErrors, field: string, index: number) {
  const globalIdx = (currentPage.value - 1) * pageSize + index
  if (!row.valid) {
    emit('edit-cell', globalIdx, field)
  }
}
</script>

<style scoped>
.excel-preview {
  margin-bottom: 16px;
}

.preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}

.preview-title { font-weight: 600; font-size: 15px; }
.preview-count { font-size: 13px; color: #909399; }

.cell-wrapper {
  display: flex;
  align-items: center;
  gap: 4px;
  min-height: 24px;
}

.cell-wrapper.cell-error { cursor: pointer; }
.cell-wrapper.cell-required::after {
  content: '*';
  color: #e74c3c;
  margin-left: 2px;
}

.error-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  background: #dc2626;
  color: white;
  font-size: 10px;
  font-weight: bold;
  cursor: help;
}

.pagination-wrapper {
  display: flex;
  justify-content: center;
  margin-top: 12px;
}
</style>
```

- [ ] **Step 2: Verify LSP diagnostics clean**

Expected: No errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add ExcelPreview component with error highlighting"
```

---

### Task 10: ExcelCellEdit.vue Component

**Files:**
- Create: `src/components/ExcelCellEdit.vue`

- [ ] **Step 1: Create component**

```vue
<template>
  <el-dialog
    :model-value="visible"
    title="编辑单元格"
    width="500px"
    @update:model-value="emit('close')"
    @close="emit('close')"
  >
    <div class="edit-info" v-if="rowData !== null">
      <p><strong>行 {{ rowIndex + 1 }}</strong> · 字段: {{ field }}</p>
      <p class="edit-errors" v-if="currentErrors.length">
        <span v-for="err in currentErrors" :key="err.message" class="edit-error-item">
          ⚠ {{ err.message }}
        </span>
      </p>
    </div>

    <el-input
      v-model="editValue"
      type="textarea"
      :rows="3"
      placeholder="输入新值"
      @keydown.esc="emit('close')"
    />

    <template #footer>
      <el-button @click="emit('close')">取消</el-button>
      <el-button type="primary" @click="save">保存</el-button>
    </template>
  </el-dialog>
</template>

<script setup lang="ts">
import { ref, watch } from 'vue'
import type { CellError } from '../types'

const props = defineProps<{
  visible: boolean
  rowIndex: number
  field: string
  value: any
  currentErrors: CellError[]
  rowData: Record<string, any> | null
}>()

const emit = defineEmits<{
  (e: 'save', rowIndex: number, field: string, value: any): void
  (e: 'close'): void
}>()

const editValue = ref('')

watch(() => props.visible, (v) => {
  if (v) editValue.value = String(props.value ?? '')
})

function save() {
  emit('save', props.rowIndex, props.field, editValue.value)
}
</script>

<style scoped>
.edit-info { margin-bottom: 12px; }
.edit-errors { margin: 4px 0; }
.edit-error-item {
  display: block;
  color: #dc2626;
  font-size: 13px;
}
</style>
```

- [ ] **Step 2: Verify LSP diagnostics clean**

Expected: No errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add ExcelCellEdit dialog component"
```

---

### Task 11: SubmitBar.vue Component

**Files:**
- Create: `src/components/SubmitBar.vue`

- [ ] **Step 1: Create component**

```vue
<template>
  <div class="submit-bar" v-if="total > 0">
    <div class="stats">
      <span class="stat-item">总行数: <strong>{{ total }}</strong></span>
      <span class="stat-divider">|</span>
      <span class="stat-item stat-pass">通过: <strong>{{ validCount }}</strong></span>
      <span class="stat-divider">|</span>
      <span class="stat-item" :class="{ 'stat-fail': errorCount > 0 }">
        错误: <strong>{{ errorCount }}</strong>
      </span>
    </div>

    <div class="actions">
      <el-progress
        v-if="submitting"
        :percentage="progressPercent"
        :stroke-width="16"
        :text-inside="true"
        style="width: 200px; margin-right: 12px;"
      />

      <el-button
        type="primary"
        size="large"
        :disabled="!canSubmit"
        :loading="submitting"
        @click="emit('submit')"
      >
        {{ submitting ? `提交中 (${progress.current}/${progress.total})` : submitText }}
      </el-button>

      <el-button
        v-if="errorCount > 0 && !submitting"
        size="large"
        @click="emit('fix-errors')"
      >
        修正后提交
      </el-button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { SubmitProgress } from '../types'

const props = defineProps<{
  total: number
  validCount: number
  errorCount: number
  submitting: boolean
  progress: SubmitProgress
}>()

const emit = defineEmits<{
  (e: 'submit'): void
  (e: 'fix-errors'): void
}>()

const canSubmit = computed(() => props.errorCount === 0 && props.total > 0 && !props.submitting)

const submitText = computed(() => {
  if (props.total === 0) return '导入数据'
  const batches = Math.ceil(props.total / 500)
  return `导入数据（${props.total}条，分${batches}批）`
})

const progressPercent = computed(() => {
  if (props.progress.total === 0) return 0
  return Math.round((props.progress.current / props.progress.total) * 100)
})
</script>

<style scoped>
.submit-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  background: #fafafa;
  border: 1px solid #e4e7ed;
  border-radius: 8px;
  flex-wrap: wrap;
  gap: 12px;
}

.stats { display: flex; align-items: center; gap: 8px; font-size: 14px; }
.stat-divider { color: #dcdfe6; }
.stat-pass strong { color: #67c23a; }
.stat-fail strong { color: #f56c6c; }

.actions { display: flex; align-items: center; }
</style>
```

- [ ] **Step 2: Verify LSP diagnostics clean**

Expected: No errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add SubmitBar component with stats and progress"
```

---

### Task 12: ExcelImport.vue Main Component

**Files:**
- Create: `src/components/ExcelImport.vue`

- [ ] **Step 1: Create main component**

```vue
<template>
  <div class="excel-import">
    <ExcelUploader
      :show-template-button="showTemplateButton"
      @file-selected="onFileSelected"
      @file-cleared="onFileCleared"
      @download-template="onDownloadTemplate"
    />

    <ExcelPreview
      v-if="validationResult.rows.length > 0"
      :columns="columns"
      :rows="validationResult.rows"
      :parse-errors="parseErrors"
      @edit-cell="openCellEdit"
    />

    <ExcelCellEdit
      :visible="editDialogVisible"
      :row-index="editRowIndex"
      :field="editField"
      :value="editValue"
      :current-errors="editErrors"
      :row-data="editRowData"
      @save="onCellSave"
      @close="editDialogVisible = false"
    />

    <SubmitBar
      :total="validationResult.total"
      :valid-count="validationResult.validCount"
      :error-count="validationResult.errorCount"
      :submitting="submitting"
      :progress="submitProgress"
      @submit="onSubmit"
    />

    <!-- Error summary -->
    <el-alert
      v-if="submitError"
      :title="submitError"
      type="error"
      show-icon
      closable
      @close="submitError = ''"
      style="margin-top: 12px"
    />

    <el-alert
      v-if="successMessage"
      :title="successMessage"
      type="success"
      show-icon
      closable
      @close="successMessage = ''"
      style="margin-top: 12px"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, reactive } from 'vue'
import { ElMessage } from 'element-plus'
import type {
  ColumnConfig, ParseResult, ValidationResult,
  RowWithErrors, CellError, SubmitProgress, SubmitResult
} from '../types'
import { useExcelParser } from '../composables/useExcelParser'
import { useExcelValidator } from '../composables/useExcelValidator'
import { useExcelSubmitter } from '../composables/useExcelSubmitter'
import { useExcelTemplate, downloadTemplateBlob } from '../composables/useExcelTemplate'
import ExcelUploader from './ExcelUploader.vue'
import ExcelPreview from './ExcelPreview.vue'
import ExcelCellEdit from './ExcelCellEdit.vue'
import SubmitBar from './SubmitBar.vue'

const props = withDefaults(defineProps<{
  columns: ColumnConfig[]
  submitApi: string | ((data: Record<string, any>[]) => Promise<any>)
  rollbackApi?: string | ((batchId: string) => Promise<any>)
  batchSize?: number
  maxFileSize?: number
  showTemplateButton?: boolean
  templateFileName?: string
  templateSheetName?: string
  templateDescription?: string
  templateExampleRows?: number
}>(), {
  batchSize: 500,
  maxFileSize: 50,
  showTemplateButton: true,
  templateFileName: '导入模板.xlsx',
  templateSheetName: 'Sheet1',
  templateExampleRows: 0
})

const emit = defineEmits<{
  (e: 'parsed', result: ParseResult): void
  (e: 'progress', progress: SubmitProgress): void
  (e: 'success', result: SubmitResult): void
  (e: 'error', error: Error): void
}>()

// State
const submitting = ref(false)
const parseErrors = ref<CellError[]>([])
const validationResult = reactive<ValidationResult>({
  valid: true, total: 0, validCount: 0, errorCount: 0, errors: [], rows: []
})
const submitProgress = reactive<SubmitProgress>({ current: 0, total: 0, batch: 0, totalBatches: 0 })
const submitError = ref('')
const successMessage = ref('')

// Cell edit state
const editDialogVisible = ref(false)
const editRowIndex = ref(0)
const editField = ref('')
const editValue = ref('')
const editErrors = ref<CellError[]>([])
const editRowData = ref<Record<string, any> | null>(null)

let currentFile: File | null = null
let currentRows: Record<string, any>[] = []
let currentParseErrors: CellError[] = []

async function onFileSelected(file: File) {
  // Check file size
  const maxBytes = props.maxFileSize! * 1024 * 1024
  if (file.size > maxBytes) {
    ElMessage.warning(`文件大小不能超过 ${props.maxFileSize}MB`)
    return
  }

  currentFile = file

  try {
    const parseResult = await useExcelParser(file, props.columns)
    currentRows = parseResult.rows
    currentParseErrors = parseResult.parseErrors
    parseErrors.value = parseResult.parseErrors

    // Validate
    const vResult = useExcelValidator(parseResult.rows, props.columns)
    Object.assign(validationResult, vResult)

    emit('parsed', parseResult)
  } catch (err: any) {
    ElMessage.error(`解析失败: ${err.message}`)
  }
}

function onFileCleared() {
  currentFile = null
  currentRows = []
  currentParseErrors = []
  parseErrors.value = []
  validationResult.total = 0
  validationResult.validCount = 0
  validationResult.errorCount = 0
  validationResult.valid = true
  validationResult.errors = []
  validationResult.rows = []
  successMessage.value = ''
  submitError.value = ''
}

async function onDownloadTemplate() {
  try {
    const blob = await useExcelTemplate(props.columns, {
      sheetName: props.templateSheetName,
      description: props.templateDescription,
      exampleRows: props.templateExampleRows
    })
    downloadTemplateBlob(blob, props.templateFileName)
  } catch (err: any) {
    ElMessage.error(`模板生成失败: ${err.message}`)
  }
}

function openCellEdit(rowIndex: number, field: string) {
  const row = validationResult.rows[rowIndex]
  if (!row) return
  editRowIndex.value = rowIndex
  editField.value = field
  editValue.value = row.data[field] ?? ''
  editErrors.value = row.errors.filter(e => e.field === field)
  editRowData.value = row.data
  editDialogVisible.value = true
}

function onCellSave(rowIndex: number, field: string, value: any) {
  const row = validationResult.rows[rowIndex]
  if (!row) return
  row.data[field] = value

  // Re-validate this row
  const rowConfig = props.columns.find(c => c.field === field)
  if (rowConfig) {
    const singleResult = useExcelValidator([row.data], props.columns)
    // Update errors for this row
    row.errors = singleResult.rows[0]?.errors ?? []
    row.valid = row.errors.length === 0
  }

  // Recalculate totals
  const validCount = validationResult.rows.filter(r => r.valid).length
  validationResult.validCount = validCount
  validationResult.errorCount = validationResult.total - validCount
  validationResult.valid = validCount === validationResult.total
  validationResult.errors = validationResult.rows.flatMap(r => r.errors)

  editDialogVisible.value = false
}

async function onSubmit() {
  const validRows = validationResult.rows
    .filter(r => r.valid)
    .map(r => r.data)

  if (validRows.length === 0) return

  submitting.value = true
  submitError.value = ''
  successMessage.value = ''

  try {
    const result = await useExcelSubmitter(validRows, {
      submitApi: props.submitApi,
      rollbackApi: props.rollbackApi,
      batchSize: props.batchSize,
      onProgress: (p) => {
        Object.assign(submitProgress, p)
        emit('progress', p)
      }
    })

    if (result.success) {
      successMessage.value = '数据导入成功！'
      emit('success', result)
    } else {
      const msg = result.rollbackSuccess
        ? `第 ${result.failedBatch} 批提交失败，已回滚`
        : `第 ${result.failedBatch} 批提交失败，回滚也失败: ${result.error}`
      submitError.value = msg
      emit('error', new Error(msg))
    }
  } catch (err: any) {
    submitError.value = `提交异常: ${err.message}`
    emit('error', err)
  } finally {
    submitting.value = false
  }
}
</script>

<style scoped>
.excel-import {
  max-width: 960px;
  margin: 0 auto;
}
</style>
```

- [ ] **Step 2: Verify LSP diagnostics clean**

Expected: No errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add ExcelImport main component"
```

---

### Task 13: Entry Point

**Files:**
- Create: `src/index.ts`

- [ ] **Step 1: Create entry point**

```typescript
import ExcelImport from './components/ExcelImport.vue'
export * from './types'
export { ExcelImport }
export default ExcelImport
```

- [ ] **Step 2: Verify build compiles**

Run: `npx vue-tsc --noEmit`
Expected: No type errors.

- [ ] **Step 3: Commit**

```bash
git add -A && git commit -m "feat: add package entry point"
```

---

### Task 14: Build & Verify

**Files:**
- (none — build artifact only)

- [ ] **Step 1: Run full test suite**

Run: `npm test`
Expected: All tests pass.

- [ ] **Step 2: Build library**

Run: `npm run build`
Expected: Produces `dist/` with ESM, UMD, and type declarations.

- [ ] **Step 3: Verify dist output**

Run: `ls dist/`
Expected: Contains `vue-excel-import.js`, `vue-excel-import.umd.cjs`, `index.d.ts`, `style.css`

- [ ] **Step 4: Commit final build config**

```bash
git add -A && git commit -m "chore: finalize build configuration"
```

---

## Spec Coverage Check

| Spec Section | Covered By |
|---|---|
| Props: columns | Task 2 (types), Task 12 (ExcelImport.vue) |
| Props: submitApi/rollbackApi/batchSize | Task 6 (useExcelSubmitter), Task 12 |
| ValidatorConfig (6 types) | Task 5 (useExcelValidator) |
| Events: @parsed/@progress/@success/@error | Task 12 (emit calls) |
| Template download | Task 7 (useExcelTemplate), Task 8 (ExcelUploader), Task 12 |
| Column order strategy (config order) | Task 4 (parser outputs in read order), Task 9 (preview order = columns[]) |
| Type conversion matrix | Task 3 (convertValue in excel.ts) |
| Date handling (serial/Unix/Date object) | Task 3 (excelSerialDateToJSDate + convertValue) |
| Batch submit + rollback | Task 6 (useExcelSubmitter) |
| Stream parsing (skip empty rows) | Task 4 (useExcelParser) |
| Page preview (100 rows/page) | Task 9 (ExcelPreview) |
| Error highlight in preview | Task 9 (cellStyle) |
| Cell inline edit | Task 10 (ExcelCellEdit), Task 12 (onCellSave) |
| Submit disabled when errors exist | Task 11 (SubmitBar: canSubmit), Task 12 |
| Package build (Vite lib mode) | Task 1 (vite.config.ts), Task 14 |
