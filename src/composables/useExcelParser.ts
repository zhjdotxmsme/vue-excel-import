import { ref, readonly, type Ref } from 'vue'
import type { ColumnConfig, ParseResult, CellError, CellValue } from '../types'
import { detectHeaderTree, matchColumnTree, type MergeRange, type MatchResult } from '../utils/header-tree'
import { setNested } from '../utils/column'

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

export interface ParserProgress {
  current: number
  total: number
  percent: number
}

export interface UseExcelParserReturn {
  /** Start parsing a file. Returns when all chunks are received. */
  parse: (file: File, columns: ColumnConfig[], maxRows?: number) => Promise<ParseResult>
  /** Reactive parsing progress (0-100%) */
  progress: Readonly<Ref<number>>
  /** Whether a parse is currently in progress */
  isLoading: Readonly<Ref<boolean>>
  /** Raw file buffer, for reuse */
  getBuffer: () => Promise<ArrayBuffer> | null
}

export function useExcelParser(): UseExcelParserReturn {
  const progress = ref(0)
  const isLoading = ref(false)
  let worker: Worker | null = null
  let bufferPromise: Promise<ArrayBuffer> | null = null

  async function parse(
    file: File,
    columns: ColumnConfig[],
    maxRows?: number
  ): Promise<ParseResult> {
    // Reset state
    progress.value = 0
    isLoading.value = true

    // Read file buffer (can be reused by submitter later)
    bufferPromise = fileToArrayBuffer(file)

    return new Promise((resolve, reject) => {
      try {
        worker = new Worker(
          new URL('./excelParser.worker.ts', import.meta.url),
          { type: 'module' }
        )
      } catch (e) {
        // Web Worker not supported — fall back to main-thread parsing
        resolve(fallbackParse(file, columns, maxRows))
        return
      }

      const accumulatedRows: Record<string, any>[] = []
      const accumulatedErrors: CellError[] = []
      let resultHeaders: string[] = []

      worker.onmessage = (e) => {
        const msg = e.data

        switch (msg.type) {
          case 'chunk':
            accumulatedRows.push(...msg.data)
            accumulatedErrors.push(...(msg.errors ?? []))
            // Only set headers from first chunk
            if (resultHeaders.length === 0 && msg.headers) {
              resultHeaders = msg.headers
            }
            progress.value = msg.progress ?? 0
            break

          case 'done': {
            isLoading.value = false
            progress.value = 100
            worker?.terminate()
            worker = null
            resolve({
              headers: resultHeaders,
              rows: accumulatedRows,
              totalRows: msg.totalRows ?? accumulatedRows.length,
              parseErrors: accumulatedErrors,
              sheets: msg.sheets ?? [],
              missingColumns: msg.missingColumns,
              unmatchedHeaders: msg.unmatchedHeaders
            })
            break
          }

          case 'error':
            isLoading.value = false
            worker?.terminate()
            worker = null
            reject(new Error(msg.message ?? '解析失败'))
            break
        }
      }

      worker.onerror = (err) => {
        isLoading.value = false
        worker?.terminate()
        worker = null
        reject(new Error(err.message ?? 'Worker 运行时错误'))
      }

      bufferPromise!.then((buffer) => {
        worker?.postMessage(
          { type: 'parse', buffer, columns, maxRows, batchSize: 500 },
          [buffer]
        )
      })
    })
  }

  function getBuffer(): Promise<ArrayBuffer> | null {
    return bufferPromise
  }

  return {
    parse,
    progress: readonly(progress),
    isLoading: readonly(isLoading),
    getBuffer
  }
}

/**
 * Fallback for environments where Web Workers are unavailable.
 * Parses on the main thread using hucre (much faster than ExcelJS).
 */
async function fallbackParse(
  file: File,
  columns: ColumnConfig[],
  maxRows?: number
): Promise<ParseResult> {
  const { readXlsx } = await import('hucre/xlsx')
  const { convertValue } = await import('../utils/excel')

  const buffer = await fileToArrayBuffer(file)
  const workbook = await readXlsx(buffer)

  if (workbook.sheets.length === 0) {
    return {
      headers: [],
      rows: [],
      totalRows: 0,
      parseErrors: [{ row: 0, field: '', value: null, message: '工作表中没有数据', type: 'parse' }],
      sheets: []
    }
  }

  const errors: CellError[] = []
  const allRows: Record<string, any>[] = []
  const sheetNames = workbook.sheets.map(s => s.name)
  let headers: string[] = []
  const maxLimit = maxRows ?? Infinity
  let firstSheet = true

  const matchedLabels = new Set<string>()
  let firstMatchResult: MatchResult | null = null

  for (const sheet of workbook.sheets) {
    const rawRows = sheet.rows
    if (rawRows.length === 0) continue

    // Detect header tree from merged cells
    const excelTree = detectHeaderTree(rawRows, { merges: (sheet as any).merges })
    // Strip trailing * markers from tree labels for matching (template compatibility)
    const cleanTree = excelTree.map(stripAsteriskNode)

    // Match against user config
    const matchResult = matchColumnTree(cleanTree, columns)
    const headerIndexMap = matchResult.columnMap

    // Save first match result for diagnostics
    if (!firstMatchResult) {
      firstMatchResult = matchResult
    }

    // Collect leaf headers (first sheet only)
    if (firstSheet) {
      headers = collectLeafLabels(cleanTree)
    }
    firstSheet = false

    // Track matched labels for diagnostics
    for (const [, cfg] of headerIndexMap) {
      matchedLabels.add(cfg.label)
    }

    // Determine header row count to know where data starts
    const headerRowCount = detectHeaderRowCount(rawRows, (sheet as any).merges)

    // Process data rows (skip header index 0)
    for (let rowIdx = headerRowCount; rowIdx < rawRows.length; rowIdx++) {
      if (allRows.length >= maxLimit) break

      const row = rawRows[rowIdx]
      if (!row) continue

      const rowData: Record<string, any> = {}
      let hasContent = false

      headerIndexMap.forEach((config, colIdx) => {
        const cellValue = row[colIdx]
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          hasContent = true
        }
        const result = convertValue(cellValue, config.type ?? 'string')
        setNested(rowData, config.field!, result.value)
        if (result.error) {
          errors.push({
            row: allRows.length + 1,
            field: config.field!,
            value: cellValue,
            message: result.error,
            type: 'type-conversion'
          })
        }
      })

      if (hasContent) allRows.push(rowData)
    }
  }

  const missingColumns = columns
    .filter(c => !matchedLabels.has(c.label))
    .map(c => c.label)
  const unmatchedHeaders = firstMatchResult?.unmatchedHeaders ?? []

  return {
    headers, rows: allRows, totalRows: allRows.length, parseErrors: errors, sheets: sheetNames,
    missingColumns, unmatchedHeaders
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

function stripAsteriskNode(n: ColumnConfig): ColumnConfig {
  return {
    ...n,
    label: n.label.replace(/\s*\*+\s*$/, ''),
    children: n.children ? n.children.map(stripAsteriskNode) : undefined
  }
}

function detectHeaderRowCount(rows: CellValue[][], merges?: MergeRange[]): number {
  if (!merges || merges.length === 0) return 1
  let maxRow = 0
  for (const m of merges) {
    if (m.endRow + 1 > maxRow) maxRow = m.endRow + 1
  }
  return Math.min(maxRow, rows.length)
}
