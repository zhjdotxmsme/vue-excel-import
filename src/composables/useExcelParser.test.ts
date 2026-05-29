import { describe, it, expect, beforeAll } from 'vitest'
import ExcelJS from 'exceljs'
import type { ColumnConfig } from '../types'
import { useExcelParser } from './useExcelParser'

// Helper: create an in-memory Excel file from headers + rows (multi-sheet support)
async function makeExcelFile(
  sheets: { name: string; headers: string[]; rows: any[][] }[]
): Promise<File>
async function makeExcelFile(headers: string[], rows: any[][]): Promise<File>
async function makeExcelFile(
  arg1: string[] | { name: string; headers: string[]; rows: any[][] }[],
  arg2?: any[][]
): Promise<File> {
  const wb = new ExcelJS.Workbook()

  if (Array.isArray(arg1) && typeof arg1[0] === 'string') {
    // Legacy: single sheet
    const ws = wb.addWorksheet('Sheet1')
    ws.addRow(arg1)
    ;(arg2 ?? []).forEach(r => ws.addRow(r))
  } else {
    // Multi-sheet
    for (const sheet of arg1 as { name: string; headers: string[]; rows: any[][] }[]) {
      const ws = wb.addWorksheet(sheet.name)
      ws.addRow(sheet.headers)
      sheet.rows.forEach(r => ws.addRow(r))
    }
  }

  const buf = await wb.xlsx.writeBuffer()
  return new File([buf], 'test.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

describe('useExcelParser', () => {
  const columns: ColumnConfig[] = [
    { label: '姓名', field: 'userName', type: 'string' },
    { label: '年龄', field: 'age', type: 'number' },
    { label: '入职日期', field: 'hireDate', type: 'date' }
  ]

  const parser = useExcelParser()

  it('parses valid Excel and maps columns by label', async () => {
    const file = await makeExcelFile(
      ['姓名', '年龄', '入职日期'],
      [['张三', 28, new Date(2024, 0, 15)], ['李四', 35, new Date(2023, 5, 1)]]
    )
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(2)
    expect(result.rows[0].userName).toBe('张三')
    expect(result.rows[0].age).toBe(28)
    expect(result.rows[0].hireDate).toBe('2024-01-15')
    expect(result.rows[1].userName).toBe('李四')
    expect(result.rows[1].age).toBe(35)
    expect(result.rows[1].hireDate).toBe('2023-06-01')
    expect(result.parseErrors).toHaveLength(0)
  })

  it('skips completely empty rows', async () => {
    const file = await makeExcelFile(
      ['姓名', '年龄'],
      [['张三', 28], [], ['李四', 35]]
    )
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(2)
  })

  it('ignores Excel columns not in column config', async () => {
    const file = await makeExcelFile(
      ['姓名', '年龄', '备注'],
      [['张三', 28, 'extra info']]
    )
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(1)
    expect(Object.keys(result.rows[0])).toEqual(['userName', 'age'])
  })

  it('handles missing columns in Excel (not found by label)', async () => {
    const file = await makeExcelFile(
      ['姓名', '年龄'],
      [['张三', 28]]
    )
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(1)
    // hireDate column not in Excel but in config — value should be undefined
    expect(result.rows[0].hireDate).toBeUndefined()
  })

  it('collects parse errors for type conversion failures', async () => {
    const file = await makeExcelFile(
      ['姓名', '年龄'],
      [['张三', 'not-a-number']]
    )
    const result = await parser.parse(file, columns)
    expect(result.parseErrors.length).toBeGreaterThan(0)
    expect(result.parseErrors[0].type).toBe('type-conversion')
    expect(result.parseErrors[0].field).toBe('age')
  })

  it('handles empty Excel file (header only, no data rows)', async () => {
    const file = await makeExcelFile(['姓名', '年龄'], [])
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(0)
    expect(result.headers).toEqual(['姓名', '年龄'])
  })

  it('returns parsed headers from Excel', async () => {
    const file = await makeExcelFile(['A', 'B', 'C'], [['1', '2', '3']])
    const result = await parser.parse(file, columns)
    expect(result.headers).toEqual(['A', 'B', 'C'])
  })

  it('strips * marker from headers when matching column configs (template compatibility)', async () => {
    // Template adds * for required columns, e.g. "姓名*" instead of "姓名"
    const file = await makeExcelFile(
      ['姓名*', '年龄', '入职日期*'],
      [['张三', 28, new Date(2024, 0, 15)]]
    )
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(1)
    expect(result.rows[0].userName).toBe('张三')
    expect(result.parseErrors).toHaveLength(0)
  })

  it('parses multiple sheets and merges rows', async () => {
    const file = await makeExcelFile([
      { name: '员工A组', headers: ['姓名', '年龄'], rows: [['张三', 28], ['李四', 35]] },
      { name: '员工B组', headers: ['姓名', '年龄'], rows: [['王五', 42]] }
    ])
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(3)
    expect(result.sheets).toEqual(['员工A组', '员工B组'])
    expect(result.rows[0].userName).toBe('张三')
    expect(result.rows[2].userName).toBe('王五')
    expect(result.parseErrors).toHaveLength(0)
  })

  it('parses sheets with same header structure', async () => {
    const file = await makeExcelFile([
      { name: 'Sheet1', headers: ['姓名', '年龄'], rows: [['张三', 28]] },
      { name: 'Sheet2', headers: ['姓名', '年龄'], rows: [['李四', 35]] }
    ])
    const result = await parser.parse(file, columns)
    expect(result.totalRows).toBe(2)
    expect(result.sheets).toEqual(['Sheet1', 'Sheet2'])
  })
})
