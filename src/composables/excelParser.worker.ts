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

    if (workbook.worksheets.length === 0) {
      self.postMessage({ type: 'error', message: '工作表中没有数据' })
      return
    }

    // Calculate total row count across all sheets for progress tracking
    const sheetNames: string[] = []
    let totalSheetRows = 0
    for (const ws of workbook.worksheets) {
      if (ws) {
        sheetNames.push(ws.name)
        // rowCount includes header row, subtract 1 for data rows
        if (ws.rowCount > 1) totalSheetRows += ws.rowCount - 1
      }
    }
    const maxLimit = maxRows ?? Infinity
    const capRows = Math.min(totalSheetRows, maxLimit)

    const errors: CellError[] = []
    let batch: Record<string, any>[] = []
    let globalRowIndex = 0
    let firstSheet = true
    let headers: string[] = []

    // Iterate all worksheets
    for (const worksheet of workbook.worksheets) {
      if (!worksheet) continue

      const headerIndexMap = new Map<number, ColumnConfig>()

      // Parse header row from this sheet
      const headerRow = worksheet.getRow(1)
      headerRow.eachCell((cell, colIdx) => {
        const rawLabel = String(cell.value ?? '').trim()
        if (firstSheet) headers.push(rawLabel)
        const cleanLabel = rawLabel.replace(/\s*\*+\s*$/, '')
        const config = columns.find(c => c.label === cleanLabel)
        if (config) headerIndexMap.set(colIdx, config)
      })
      firstSheet = false

      // Only iterate rows that have content
      worksheet.eachRow({ includeEmpty: false }, (row, rowIdx) => {
        // Skip header row
        if (rowIdx === 1) return

        // Stop if we've hit the row limit
        if (globalRowIndex >= maxLimit) {
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
              row: globalRowIndex + 1,
              field: config.field,
              value: cell.value,
              message: result.error,
              type: 'type-conversion'
            })
          }
        })

        batch.push(rowData)
        globalRowIndex++

        // Send chunk when batch is full
        if (batch.length >= batchSize) flush()
      })
    }

    // Send remaining data
    if (batch.length > 0) flush()

    // Signal completion with sheet info
    self.postMessage({ type: 'done', totalRows: globalRowIndex, totalErrors: errors.length, sheets: sheetNames })

    function flush() {
      self.postMessage({
        type: 'chunk',
        data: batch,
        headers,
        errors: [...errors],
        progress: capRows > 0 ? Math.min(Math.round((globalRowIndex / capRows) * 100), 100) : 100,
        totalRows: globalRowIndex
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
