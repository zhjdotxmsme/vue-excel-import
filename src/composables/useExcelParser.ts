import ExcelJS from 'exceljs'
import type { ColumnConfig, ParseResult, CellError } from '../types'
import { convertValue } from '../utils/excel'

/**
 * Read a File/Blob as ArrayBuffer.
 * Falls back to FileReader when file.arrayBuffer() is unavailable (e.g., jsdom).
 */
async function fileToArrayBuffer(file: File): Promise<ArrayBuffer> {
  if (typeof file.arrayBuffer === 'function') {
    return file.arrayBuffer()
  }
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as ArrayBuffer)
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

export async function useExcelParser(
  file: File,
  columns: ColumnConfig[]
): Promise<ParseResult> {
  const workbook = new ExcelJS.Workbook()
  const buffer = await fileToArrayBuffer(file)
  await workbook.xlsx.load(buffer)

  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    return {
      headers: [],
      rows: [],
      totalRows: 0,
      parseErrors: [{ row: 0, field: '', value: null, message: '工作表中没有数据', type: 'parse' }]
    }
  }

  const errors: CellError[] = []
  const rows: Record<string, any>[] = []
  let headers: string[] = []
  const headerIndexMap = new Map<number, ColumnConfig>()

  // Read header row (row 1) and map column indices to configs
  const headerRow = worksheet.getRow(1)
  headerRow.eachCell((cell, colIdx) => {
    const rawLabel = String(cell.value ?? '').trim()
    headers.push(rawLabel)
    // Strip * markers (from template download) before matching against column labels
    const cleanLabel = rawLabel.replace(/\s*\*+\s*$/, '')
    const config = columns.find(c => c.label === cleanLabel)
    if (config) headerIndexMap.set(colIdx, config)
  })

  // Iterate data rows (starting from row 2)
  worksheet.eachRow((row, rowIdx) => {
    if (rowIdx === 1) return

    let hasContent = false
    const rowData: Record<string, any> = {}

    row.eachCell((cell, colIdx) => {
      const isEmpty = cell.value === null || cell.value === undefined || cell.value === ''
      if (!isEmpty) hasContent = true

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
