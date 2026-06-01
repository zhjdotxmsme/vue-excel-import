import { ref, readonly, type Ref } from 'vue'
import type { ColumnConfig, ParseResult, CellError } from '../types'

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

  for (const sheet of workbook.sheets) {
    const headerIndexMap = new Map<number, ColumnConfig>()
    const rawRows = sheet.rows
    if (rawRows.length === 0) continue

    // First row is the header
    const headerRow = rawRows[0]
    if (headerRow) {
      headerRow.forEach((cellValue, colIdx) => {
        const label = String(cellValue ?? '').trim()
        if (firstSheet) headers.push(label)
        const cleanLabel = label.replace(/\s*\*+\s*$/, '')
        const config = columns.find(c => c.label === cleanLabel)
        if (config) {
          headerIndexMap.set(colIdx, config)
          matchedLabels.add(config.label)
        }
      })
    }
    firstSheet = false

    // Process data rows (skip header index 0)
    for (let rowIdx = 1; rowIdx < rawRows.length; rowIdx++) {
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
        rowData[config.field] = result.value
        if (result.error) {
          errors.push({
            row: allRows.length + 1,
            field: config.field,
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
  const unmatchedHeaders = headers.filter(h => {
    const cleaned = h.replace(/\s*\*+\s*$/, '')
    return !columns.some(c => c.label === cleaned)
  })

  return {
    headers, rows: allRows, totalRows: allRows.length, parseErrors: errors, sheets: sheetNames,
    missingColumns, unmatchedHeaders
  }
}
