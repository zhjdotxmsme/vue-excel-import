import ExcelJS from 'exceljs'
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
  const workbook = new ExcelJS.Workbook()
  const sheetName = options?.sheetName ?? 'Sheet1'
  const ws = workbook.addWorksheet(sheetName)

  // Optional description row
  let headerRowIndex = 1
  if (options?.description) {
    ws.addRow([options.description])
    headerRowIndex = 2
  }

  // Header row
  const headerLabels = columns.map(col => {
    return col.required ? `${col.label} *` : col.label
  })
  ws.addRow(headerLabels)

  // Style header row
  const headerExcelRow = ws.getRow(headerRowIndex)
  headerExcelRow.font = { bold: true }
  headerExcelRow.height = 28

  // Optional example rows
  if (options?.exampleRows && options.exampleRows > 0) {
    const now = new Date()
    for (let i = 0; i < options.exampleRows; i++) {
      const exampleRow = columns.map(col => {
        switch (col.type) {
          case 'number': return 0
          case 'date': return now
          case 'boolean': return true
          default: return ''
        }
      })
      ws.addRow(exampleRow)
    }
  }

  // Set column widths and formats
  columns.forEach((col, idx) => {
    const colIdx = idx + 1
    ws.getColumn(colIdx).width = 18

    switch (col.type) {
      case 'number':
        ws.getColumn(colIdx).numFmt = '0'
        break
      case 'date':
        ws.getColumn(colIdx).numFmt = 'yyyy-mm-dd'
        break
    }

    // Data validation for enum columns
    const enumValidator = col.validators?.find(v => v.type === 'enum')
    if (enumValidator?.enum && enumValidator.enum.length > 0) {
      const colLetter = String.fromCharCode(64 + colIdx)
      const firstRow = headerRowIndex + 1
      const lastRow = Math.max(firstRow, ws.rowCount)
      ;(ws as any).dataValidations.add(`${colLetter}${firstRow}:${colLetter}${lastRow}`, {
        type: 'list',
        formulae: [enumValidator.enum.join(',')],
        showErrorMessage: true,
        errorTitle: '输入错误',
        error: `请选择: ${enumValidator.enum.join(', ')}`
      })
    }
  })

  // Freeze header row
  ws.views = [{ state: 'frozen', ySplit: headerRowIndex }]

  // Generate blob
  const buffer = await workbook.xlsx.writeBuffer()
  return new Blob([buffer], {
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
