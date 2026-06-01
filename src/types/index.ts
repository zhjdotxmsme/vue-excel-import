export type ColumnType = 'string' | 'number' | 'boolean' | 'date'

export type CellValue = string | number | boolean | Date | null

export interface ColumnConfig {
  label: string
  field?: string
  type?: ColumnType
  required?: boolean
  default?: any
  description?: string
  validators?: ValidatorConfig[]
  children?: ColumnConfig[]
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
  sheets: string[]
  /** 配置了但在 Excel 中未找到的列名 */
  missingColumns?: string[]
  /** Excel 中的表头但未在 columns 配置中匹配到的列 */
  unmatchedHeaders?: string[]
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