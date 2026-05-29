import { describe, it, expect } from 'vitest'
import type { ColumnConfig } from '../types'
import { useExcelValidator } from './useExcelValidator'

describe('useExcelValidator', () => {
  it('returns valid=true when no validators configured and all valid', () => {
    const columns: ColumnConfig[] = [{ label: 'Name', field: 'name' }]
    const result = useExcelValidator([{ name: 'Alice' }], columns)
    expect(result.valid).toBe(true)
    expect(result.validCount).toBe(1)
    expect(result.errorCount).toBe(0)
  })

  it('validates required fields (from ColumnConfig.required)', () => {
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
      { label: 'Email', field: 'email', validators: [{ type: 'unique' }] }
    ]
    const result = useExcelValidator(
      [{ email: 'a@b.com' }, { email: 'a@b.com' }, { email: 'c@d.com' }],
      columns
    )
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(2)
  })

  it('validates range (min/max) for numbers', () => {
    const columns: ColumnConfig[] = [
      { label: 'Age', field: 'age', type: 'number', validators: [{ type: 'range', min: 0, max: 150 }] }
    ]
    const result = useExcelValidator([{ age: 25 }, { age: -1 }, { age: 200 }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(2)
  })

  it('validates pattern (regex)', () => {
    const columns: ColumnConfig[] = [
      { label: 'Email', field: 'email', validators: [{ type: 'pattern', pattern: '^\\S+@\\S+\\.\\S+$' }] }
    ]
    const result = useExcelValidator([{ email: 'good@test.com' }, { email: 'bad' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errors).toHaveLength(1)
    expect(result.errors[0].type).toBe('pattern')
  })

  it('validates enum values', () => {
    const columns: ColumnConfig[] = [
      { label: 'Status', field: 'status', validators: [{ type: 'enum', enum: ['active', 'inactive'] }] }
    ]
    const result = useExcelValidator([{ status: 'active' }, { status: 'unknown' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
  })

  it('validates with custom sync function', () => {
    const columns: ColumnConfig[] = [{
      label: 'Code', field: 'code',
      validators: [{ type: 'custom', validate: (v) => v?.startsWith('CODE-') ? true : '必须以 CODE- 开头' }]
    }]
    const result = useExcelValidator([{ code: 'CODE-123' }, { code: 'BAD' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errors[0].message).toBe('必须以 CODE- 开头')
  })

  it('validates with custom async function', async () => {
    const columns: ColumnConfig[] = [{
      label: 'Email', field: 'email',
      validators: [{
        type: 'custom',
        validate: async (v) => {
          await new Promise(r => setTimeout(r, 5))
          return v?.includes('@') ? true : 'invalid email'
        }
      }]
    }]
    const result = await useExcelValidator([{ email: 'a@b.com' }, { email: 'abc' }], columns)
    expect(result.valid).toBe(false)
    expect(result.errorCount).toBe(1)
  })

  it('skips non-required validators for empty values', () => {
    const columns: ColumnConfig[] = [
      { label: 'Age', field: 'age', validators: [{ type: 'range', min: 0, max: 150 }] }
    ]
    const result = useExcelValidator([{ age: null }, { age: '' }], columns)
    expect(result.valid).toBe(true)
  })

  it('marks rows with RowWithErrors.valid correctly', () => {
    const columns: ColumnConfig[] = [
      { label: 'Name', field: 'name', required: true }
    ]
    const result = useExcelValidator([{ name: 'Alice' }, { name: '' }], columns)
    expect(result.rows[0].valid).toBe(true)
    expect(result.rows[1].valid).toBe(false)
    expect(result.rows[0].errors).toHaveLength(0)
    expect(result.rows[1].errors).toHaveLength(1)
  })
})
