import { readXlsx } from 'hucre/xlsx'
import type { ColumnConfig, CellError, CellValue } from '../types'
import { convertValue } from '../utils/excel'
import { detectHeaderTree, matchColumnTree, type MergeRange, type MatchResult } from '../utils/header-tree'
import { setNested } from '../utils/column'

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
    let firstMatchResult: MatchResult | null = null
    let headers: string[] = []
    const matchedLabels = new Set<string>()
    const maxLimit = maxRows ?? Infinity

    for (const sheet of workbook.sheets) {
      const rawRows = sheet.rows

      if (rawRows.length === 0) continue

      // Detect header tree from merged cells
      const excelTree = detectHeaderTree(rawRows, { merges: (sheet as any).merges })

      // Match against user config
      const matchResult = matchColumnTree(excelTree, columns)
      const headerIndexMap = matchResult.columnMap

      // Save first match result for diagnostics
      if (!firstMatchResult) {
        firstMatchResult = matchResult
      }

      // Collect leaf headers for backward-compat headers array (first sheet only)
      if (firstSheet) {
        headers = collectLeafLabels(excelTree)
      }
      firstSheet = false

      // Track matched labels for diagnostics
      for (const [, cfg] of headerIndexMap) {
        matchedLabels.add(cfg.label)
      }

      // Determine how many header rows exist (to skip them)
      const headerRowCount = detectHeaderRowCount(rawRows, (sheet as any).merges)

      // Process data rows (skip header rows)
      for (let rowIdx = headerRowCount; rowIdx < rawRows.length; rowIdx++) {
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
          setNested(rowData, config.field, result.value)
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

    // Compute column matching diagnostics
    const missingColumns = columns
      .filter(c => !matchedLabels.has(c.label))
      .map(c => c.label)
    const unmatchedHeaders = firstMatchResult?.unmatchedHeaders ?? []

    // Signal completion with sheet info and header diagnostics
    self.postMessage({
      type: 'done',
      totalRows: globalRowIndex,
      totalErrors: errors.length,
      sheets: sheetNames,
      missingColumns,
      unmatchedHeaders
    })

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

function collectLeafLabels(nodes: ColumnConfig[]): string[] {
  const result: string[] = []
  for (const n of nodes) {
    if (n.children && n.children.length > 0) {
      result.push(...collectLeafLabels(n.children))
    } else {
      result.push(n.label)
    }
  }
  return result
}

function detectHeaderRowCount(rows: CellValue[][], merges?: MergeRange[]): number {
  if (!merges || merges.length === 0) return 1
  let maxRow = 0
  for (const m of merges) {
    if (m.endRow + 1 > maxRow) maxRow = m.endRow + 1
  }
  return Math.min(maxRow, rows.length)
}
