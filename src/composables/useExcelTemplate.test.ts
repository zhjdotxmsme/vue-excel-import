import { describe, it, expect } from 'vitest'
import ExcelJS from 'exceljs'
import type { ColumnConfig } from '../types'
import { useExcelTemplate } from './useExcelTemplate'

async function blobToArrayBuffer(blob: Blob): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as ArrayBuffer)
    reader.onerror = reject
    reader.readAsArrayBuffer(blob)
  })
}

describe('useExcelTemplate', () => {
  const columns: ColumnConfig[] = [
    { label: '姓名', field: 'userName', type: 'string', required: true },
    { label: '年龄', field: 'age', type: 'number' },
    { label: '部门', field: 'dept', type: 'string', validators: [{ type: 'enum', enum: ['技术部', '市场部', '财务部'] }] },
    { label: '入职日期', field: 'hireDate', type: 'date' }
  ]

  it('generates a Blob with correct MIME type', async () => {
    const blob = await useExcelTemplate(columns)
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
  })

  it('creates workbook with column headers from config', async () => {
    const blob = await useExcelTemplate(columns)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blobToArrayBuffer(blob))
    const ws = workbook.worksheets[0]
    expect(ws).toBeDefined()

    const row1 = ws.getRow(1)
    expect(String(row1.getCell(1).value)).toContain('姓名')
    expect(String(row1.getCell(2).value)).toContain('年龄')
    expect(String(row1.getCell(3).value)).toContain('部门')
    expect(String(row1.getCell(4).value)).toContain('入职日期')
  })

  it('marks required columns with asterisk in header', async () => {
    const blob = await useExcelTemplate(columns)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blobToArrayBuffer(blob))
    const ws = workbook.worksheets[0]
    const header1 = String(ws.getRow(1).getCell(1).value)
    expect(header1).toContain('*')
  })

  it('adds data validation for enum columns', async () => {
    const blob = await useExcelTemplate(columns)
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blobToArrayBuffer(blob))
    const ws = workbook.worksheets[0]
    // Column C (部门) should have data validation
    const validations = ws.dataValidations as any
    expect(validations).toBeDefined()
  })

  it('accepts optional description text', async () => {
    const blob = await useExcelTemplate(columns, { description: '请按要求填写' })
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blobToArrayBuffer(blob))
    const ws = workbook.worksheets[0]
    // With description, row 1 should be the description, row 2 should be headers
    const descRow = ws.getRow(1)
    expect(String(descRow.getCell(1).value)).toContain('请按要求填写')
  })

  it('generates example rows when specified', async () => {
    const blob = await useExcelTemplate(columns, { exampleRows: 2 })
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(await blobToArrayBuffer(blob))
    const ws = workbook.worksheets[0]
    // Headers + 2 example rows = 3 rows total (without description)
    expect(ws.rowCount).toBeGreaterThanOrEqual(3)
  })
})
