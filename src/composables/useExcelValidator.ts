import type { ColumnConfig, ValidationResult, RowWithErrors, CellError, ValidatorConfig } from '../types'
import { getNested, flattenColumns } from '../utils/column'

function runValidator(
  value: any,
  row: Record<string, any>,
  allData: Record<string, any>[],
  config: ValidatorConfig,
  field: string,
  rowIdx: number
): CellError | null | Promise<CellError | null> {
  const msg = config.message ?? ''

  switch (config.type) {
    case 'required': {
      if (value === null || value === undefined || value === '') {
        return { row: rowIdx, field, value, message: msg || '此字段为必填项', type: 'required' }
      }
      return null
    }

    case 'unique': {
      const count = allData.filter(d => getNested(d, field) === value).length
      if (count > 1) {
        return { row: rowIdx, field, value, message: msg || `值 "${value}" 已存在`, type: 'unique' }
      }
      return null
    }

    case 'range': {
      if (value === null || value === undefined || value === '') return null
      const num = Number(value)
      if (isNaN(num)) return { row: rowIdx, field, value, message: msg || '不是有效数字', type: 'range' }
      if (config.min !== undefined && num < config.min) {
        return { row: rowIdx, field, value, message: msg || `不能小于 ${config.min}`, type: 'range' }
      }
      if (config.max !== undefined && num > config.max) {
        return { row: rowIdx, field, value, message: msg || `不能大于 ${config.max}`, type: 'range' }
      }
      return null
    }

    case 'pattern': {
      if (value === null || value === undefined || value === '') return null
      const pattern = config.pattern instanceof RegExp ? config.pattern : new RegExp(config.pattern!)
      if (!pattern.test(String(value))) {
        return { row: rowIdx, field, value, message: msg || '格式不匹配', type: 'pattern' }
      }
      return null
    }

    case 'enum': {
      if (value === null || value === undefined || value === '') return null
      if (!config.enum!.includes(value)) {
        return { row: rowIdx, field, value, message: msg || `值 "${value}" 不在允许列表中`, type: 'enum' }
      }
      return null
    }

    case 'custom': {
      const result = config.validate!(value, row, allData)
      const handleResult = (r: boolean | string): CellError | null => {
        if (r === true || r === '') return null
        return { row: rowIdx, field, value, message: typeof r === 'string' ? r : (msg || '自定义校验未通过'), type: 'custom' }
      }
      if (result instanceof Promise) return result.then(handleResult)
      return handleResult(result)
    }

    default:
      return null
  }
}

function hasAsyncValidator(columns: ColumnConfig[]): boolean {
  return columns.some(c =>
    c.validators?.some(v => v.type === 'custom' && v.validate!.constructor.name === 'AsyncFunction')
  )
}

function runValidationSync(
  rows: Record<string, any>[],
  columns: ColumnConfig[]
): ValidationResult {
  const errors: CellError[] = []

  const rowWrappers: RowWithErrors[] = rows.map((data, idx) => {
    const rowErrors: CellError[] = []
    const rowNum = idx + 1

    const leafColumns = flattenColumns(columns)
    for (const col of leafColumns) {
      const value = getNested(data, col.field!)

      // Handle required from ColumnConfig.required
      if (col.required && (value === null || value === undefined || value === '')) {
        rowErrors.push({ row: rowNum, field: col.field!, value, message: '此字段为必填项', type: 'required' })
        continue
      }

      if (col.validators) {
        for (const vc of col.validators) {
          if (vc.type === 'required') continue
          const err = runValidator(value, data, rows, vc, col.field!, rowNum)
          if (err) rowErrors.push(err as CellError)
        }
      }
    }

    errors.push(...rowErrors)
    return { data, errors: rowErrors, valid: rowErrors.length === 0 }
  })

  const validCount = rowWrappers.filter(r => r.valid).length
  return { valid: validCount === rows.length, total: rows.length, validCount, errorCount: rows.length - validCount, errors, rows: rowWrappers }
}

async function runValidationAsync(
  rows: Record<string, any>[],
  columns: ColumnConfig[]
): Promise<ValidationResult> {
  const errors: CellError[] = []

  const rowWrappers: RowWithErrors[] = await Promise.all(rows.map(async (data, idx) => {
    const rowErrors: CellError[] = []
    const rowNum = idx + 1

    const leafColumns = flattenColumns(columns)
    for (const col of leafColumns) {
      const value = getNested(data, col.field!)

      if (col.required && (value === null || value === undefined || value === '')) {
        rowErrors.push({ row: rowNum, field: col.field!, value, message: '此字段为必填项', type: 'required' })
        continue
      }

      if (col.validators) {
        for (const vc of col.validators) {
          if (vc.type === 'required') continue
          const err = await runValidator(value, data, rows, vc, col.field!, rowNum)
          if (err) rowErrors.push(err)
        }
      }
    }

    errors.push(...rowErrors)
    return { data, errors: rowErrors, valid: rowErrors.length === 0 }
  }))

  const validCount = rowWrappers.filter(r => r.valid).length
  return { valid: validCount === rows.length, total: rows.length, validCount, errorCount: rows.length - validCount, errors, rows: rowWrappers }
}

export function useExcelValidator(
  rows: Record<string, any>[],
  columns: ColumnConfig[]
): ValidationResult {
  if (hasAsyncValidator(columns)) {
    return runValidationAsync(rows, columns) as unknown as ValidationResult
  }
  return runValidationSync(rows, columns)
}
