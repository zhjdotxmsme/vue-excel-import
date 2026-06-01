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

  if (options?.description) {
    rows.push([options.description])
  }

  // Header row: bold by default via style
  const headerLabels = columns.map(col =>
    col.required ? `${col.label} *` : col.label
  )
  rows.push(headerLabels)

  // Example rows
  if (options?.exampleRows && options.exampleRows > 0) {
    const now = new Date()
    for (let i = 0; i < options.exampleRows; i++) {
      rows.push(
        columns.map(col => {
          switch (col.type) {
            case 'number': return 0
            case 'date': return now
            case 'boolean': return true
            default: return ''
          }
        })
      )
    }
  }

  // Build column definitions with widths and formatting
  const headerRowIndex = options?.description ? 2 : 1
  const sheetColumns = columns.map((col, idx) => {
    const def: any = {
      header: headerLabels[idx],
      width: 18
    }

    // Number/date formatting via column numFmt
    if (col.type === 'number') {
      def.style = { numFmt: '0' }
    } else if (col.type === 'date') {
      def.style = { numFmt: 'yyyy-mm-dd' }
    }

    return def
  })

  // Data validations for enum columns
  const dataValidations: any[] = []
  columns.forEach((col, idx) => {
    const enumValidator = col.validators?.find(v => v.type === 'enum')
    if (enumValidator?.enum && enumValidator.enum.length > 0) {
      const colLetter = String.fromCharCode(65 + idx)
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
  })

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
