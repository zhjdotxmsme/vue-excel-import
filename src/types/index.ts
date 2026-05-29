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
  sheets: string[]
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