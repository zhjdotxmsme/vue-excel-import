import { writeXlsx } from 'hucre/xlsx'
import type { ColumnConfig } from '../types'

interface TemplateOptions {
  fileName?: string
  sheetName?: string
  description?: string
  exampleRows?: number
}

export async function useExcelTemplate(
  columns: ColumnConfig[],
  options?: TemplateOptions
): Promise<Blob> {
  if (!columns || columns.length === 0) {
    throw new Error('列配置不能为空，请先配置 columns')
  }

  // Build rows: description (optional) + header + example rows
  const rows: any[][] = []
  const merges: any[] = []

  if (options?.description) {
    rows.push([options.description])
  }

  // --- Recursive header building for multi-level nested ColumnConfig ---

  function getMaxDepth(nodes: ColumnConfig[]): number {
    let max = 0
    for (const n of nodes) {
      if (n.children && n.children.length > 0) {
        const d = getMaxDepth(n.children) + 1
        if (d > max) max = d
      } else {
        if (1 > max) max = 1
      }
    }
    return max
  }

  function countLeaves(nodes: ColumnConfig[]): number {
    let count = 0
    for (const n of nodes) {
      if (n.children && n.children.length > 0) {
        count += countLeaves(n.children)
      } else {
        count++
      }
    }
    return count
  }

  const depth = getMaxDepth(columns)
  const headerRows: string[][] = Array.from({ length: depth }, () => [])

  function buildHeaderRow(
    nodes: ColumnConfig[],
    currentDepth: number,
    startCol: number
  ): number {
    let col = startCol
    for (const node of nodes) {
      if (node.children && node.children.length > 0) {
        const childCount = countLeaves(node.children)
        const endCol = col + childCount - 1
        headerRows[currentDepth][col] = node.label
        const groupDepth = getMaxDepth([node])
        if (groupDepth > 1) {
          merges.push({
            startRow: currentDepth,
            startCol: col,
            endRow: currentDepth + groupDepth - 1,
            endCol: endCol,
          })
        }
        // Fill null for spanned columns in current row
        for (let c = col + 1; c <= endCol; c++) {
          if (headerRows[currentDepth][c] == null) headerRows[currentDepth][c] = null as any
        }
        const nextCol = buildHeaderRow(node.children, currentDepth + 1, col)
        col = nextCol
      } else {
        const label = node.required ? `${node.label} *` : node.label
        headerRows[currentDepth][col] = label
        col++
      }
    }
    return col
  }

  buildHeaderRow(columns, 0, 0)

  // Push pre-built header rows
  for (const hr of headerRows) {
    rows.push(hr)
  }

  // --- Example rows (recursive, leaf-only) ---
  if (options?.exampleRows && options.exampleRows > 0) {
    const now = new Date()
    for (let i = 0; i < options.exampleRows; i++) {
      const exampleRow: any[] = []
      function buildExampleRow(nodes: ColumnConfig[]) {
        for (const node of nodes) {
          if (node.children && node.children.length > 0) {
            buildExampleRow(node.children)
          } else {
            switch (node.type) {
              case 'number': exampleRow.push(0); break
              case 'date': exampleRow.push(now); break
              case 'boolean': exampleRow.push(true); break
              default: exampleRow.push(''); break
            }
          }
        }
      }
      buildExampleRow(columns)
      rows.push(exampleRow)
    }
  }

  // --- Column definitions with widths and formatting (recursive, leaf-only) ---
  const headerRowIndex = options?.description ? depth + 1 : depth
  const sheetColumns: any[] = []
  const dataValidations: any[] = []
  let colIdx = 0

  function buildColumns(nodes: ColumnConfig[]) {
    for (const node of nodes) {
      if (node.children && node.children.length > 0) {
        buildColumns(node.children)
      } else {
        const def: any = { header: node.label, width: 18 }
        if (node.type === 'number') {
          def.style = { numFmt: '0' }
        } else if (node.type === 'date') {
          def.style = { numFmt: 'yyyy-mm-dd' }
        }
        sheetColumns.push(def)

        // Enum validations
        const enumValidator = node.validators?.find(v => v.type === 'enum')
        if (enumValidator?.enum && enumValidator.enum.length > 0) {
          const colLetter = String.fromCharCode(65 + colIdx)
          const lastRow = rows.length
          dataValidations.push({
            type: 'list',
            values: enumValidator.enum,
            ranges: [`${colLetter}${headerRowIndex}:${colLetter}${Math.max(headerRowIndex, lastRow)}`],
            showErrorMessage: true,
            errorTitle: '输入错误',
            errorMessage: `请选择: ${enumValidator.enum.join(', ')}`
          })
        }
        colIdx++
      }
    }
  }
  buildColumns(columns)

  // Build the sheet
  const sheet: any = {
    name: options?.sheetName ?? 'Sheet1',
    rows,
    freezePane: { rows: headerRowIndex }
  }

  // Add column definitions if we have any with styles
  if (sheetColumns.some((c: any) => c.style)) {
    sheet.columns = sheetColumns
  }

  // Add data validations
  if (dataValidations.length > 0) {
    sheet.dataValidations = dataValidations
  }

  // Add merges for multi-level headers
  if (merges.length > 0) {
    sheet.merges = merges
  }

  // Write to XLSX
  const uint8 = await writeXlsx({ sheets: [sheet] })

  return new Blob([uint8], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
}

export function downloadTemplateBlob(blob: Blob, fileName: string = '导入模板.xlsx'): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
