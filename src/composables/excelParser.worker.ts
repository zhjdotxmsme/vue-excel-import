import ExcelJS from 'exceljs'
import type { ColumnConfig, CellError } from '../types'
import { convertValue } from '../utils/excel'

interface ParseRequest {
  type: 'parse'
  buffer: ArrayBuffer
  columns: ColumnConfig[]
  maxRows?: number
  batchSize?: number
}

self.onmessage = async (e: MessageEvent<ParseRequest>) => {
  if (e.data.type !== 'parse') return

  const { buffer, columns, maxRows, batchSize = 500 } = e.data

  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)

    const worksheet = workbook.worksheets[0]
    if (!worksheet) {
      self.postMessage({ type: 'error', message: '工作表中没有数据' })
      return
    }

    const errors: CellError[] = []
    let headers: string[] = []
    const headerIndexMap = new Map<number, ColumnConfig>()
    let batch: Record<string, any>[] = []
    let totalRows = 0
    const maxLimit = maxRows ?? Infinity

    // Parse header row
    const headerRow = worksheet.getRow(1)
    headerRow.eachCell((cell, colIdx) => {
      const rawLabel = String(cell.value ?? '').trim()
      headers.push(rawLabel)
      const cleanLabel = rawLabel.replace(/\s*\*+\s*$/, '')
      const config = columns.find(c => c.label === cleanLabel)
      if (config) headerIndexMap.set(colIdx, config)
    })

    // Only iterate rows that have content — skip empty rows entirely
    // ExcelJS eachRow({ includeEmpty: false }) skips rows without any cell values
    worksheet.eachRow({ includeEmpty: false }, (row, rowIdx) => {
      // Skip header row
      if (rowIdx === 1) return

      // Stop if we've hit the row limit
      if (totalRows >= maxLimit) {
        if (batch.length > 0) flush()
        return false
      }

      const rowData: Record<string, any> = {}

      row.eachCell((cell, colIdx) => {
        const config = headerIndexMap.get(colIdx)
        if (!config) return

        const result = convertValue(cell.value, config.type ?? 'string')
        rowData[config.field] = result.value

        if (result.error) {
          errors.push({
            row: totalRows + 1,
            field: config.field,
            value: cell.value,
            message: result.error,
            type: 'type-conversion'
          })
        }
      })

      batch.push(rowData)
      totalRows++

      // Send chunk when batch is full
      if (batch.length >= batchSize) flush()
    })

    // Send remaining data
    if (batch.length > 0) flush()

    // Signal completion
    self.postMessage({ type: 'done', totalRows, totalErrors: errors.length })

    function flush() {
      self.postMessage({
        type: 'chunk',
        data: batch,
        headers,
        errors: [...errors],
        progress: Math.min(
          Math.round((totalRows / Math.min(worksheet.rowCount, maxLimit)) * 100),
          100
        ),
        totalRows
      })
      batch = []
      errors.length = 0
    }
  } catch (err: any) {
    self.postMessage({
      type: 'error',
      message: err?.message ?? '解析 Excel 文件时发生未知错误'
    })
  }
}
