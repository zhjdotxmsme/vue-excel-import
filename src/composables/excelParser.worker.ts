import { readXlsx } from 'hucre/xlsx'
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
    // hucre: zero-dependency minimal XLSX parser
    // Skips styles, themes, charts — only reads cell values + sheet names
    const workbook = await readXlsx(buffer)

    if (workbook.sheets.length === 0) {
      self.postMessage({ type: 'error', message: '工作表中没有数据' })
      return
    }

    const sheetNames = workbook.sheets.map(s => s.name)
    const errors: CellError[] = []
    let batch: Record<string, any>[] = []
    let globalRowIndex = 0
    let firstSheet = true
    let headers: string[] = []
    const maxLimit = maxRows ?? Infinity

    for (const sheet of workbook.sheets) {
      const headerIndexMap = new Map<number, ColumnConfig>()
      const rawRows = sheet.rows

      if (rawRows.length === 0) continue

      // First row is the header row
      const headerRow = rawRows[0]
      if (headerRow) {
        headerRow.forEach((cellValue, colIdx) => {
          const label = String(cellValue ?? '').trim()
          if (firstSheet) headers.push(label)
          const cleanLabel = label.replace(/\s*\*+\s*$/, '')
          const config = columns.find(c => c.label === cleanLabel)
          if (config) headerIndexMap.set(colIdx, config)
        })
      }
      firstSheet = false

      // Process data rows (skip header row index 0)
      for (let rowIdx = 1; rowIdx < rawRows.length; rowIdx++) {
        if (globalRowIndex >= maxLimit) break

        const row = rawRows[rowIdx]
        if (!row) continue

        // Check if row has content (at least one non-null value in a mapped column)
        let hasContent = false
        const rowData: Record<string, any> = {}

        headerIndexMap.forEach((config, colIdx) => {
          const cellValue = row[colIdx]
          if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
            hasContent = true
          }
          const result = convertValue(cellValue, config.type ?? 'string')
          rowData[config.field] = result.value
          if (result.error) {
            errors.push({
              row: globalRowIndex + 1,
              field: config.field,
              value: cellValue,
              message: result.error,
              type: 'type-conversion'
            })
          }
        })

        if (!hasContent) continue

        batch.push(rowData)
        globalRowIndex++

        // Send chunk when batch is full
        if (batch.length >= batchSize) flush()
      }
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
        progress: maxLimit > 0 ? Math.min(Math.round((globalRowIndex / maxLimit) * 100), 100) : 100,
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
