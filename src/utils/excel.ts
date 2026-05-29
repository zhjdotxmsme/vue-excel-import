import type { ColumnType } from '../types'

const EXCEL_EPOCH_DAYS = 25568

export function excelSerialDateToJSDate(serial: number): Date {
  return new Date((serial - EXCEL_EPOCH_DAYS) * 86400000)
}

export function formatDate(date: Date): string {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

// Extract a clean value from ExcelJS cell value (handles richText, error, hyperlink)
function extractCellValue(raw: any): any {
  if (raw === null || raw === undefined) return null
  if (typeof raw !== 'object') return raw
  if ('richText' in raw) return raw.richText.map((r: any) => r.text).join('')
  if ('error' in raw) return raw
  if ('hyperlink' in raw) return raw.text ?? String(raw.hyperlink)
  return raw
}

export function convertValue(raw: any, type: ColumnType): { value: any; error?: string } {
  const v = extractCellValue(raw)
  if (v === null || v === undefined) return type === 'string' ? { value: '' } : { value: null }

  // Excel error object
  if (typeof v === 'object' && v !== null && 'error' in v) {
    if (type === 'string') return { value: String(v.error) }
    return { value: null, error: `Excel error: ${v.error}` }
  }

  switch (type) {
    case 'string': {
      if (typeof v === 'string') return { value: v.trim() }
      if (typeof v === 'number') return { value: Number.isInteger(v) ? String(v) : String(v) }
      if (typeof v === 'boolean') return { value: String(v) }
      if (v instanceof Date) return { value: formatDate(v) }
      return { value: String(v) }
    }

    case 'number': {
      if (typeof v === 'number') return { value: v }
      if (typeof v === 'boolean') return { value: v ? 1 : 0 }
      if (typeof v === 'string') {
        const t = v.trim()
        if (!t) return { value: null }
        const n = Number(t)
        return isNaN(n) ? { value: v, error: `"${v}" 无法转换为数字` } : { value: n }
      }
      if (v instanceof Date) return { value: v.getTime() }
      return { value: v, error: `无法将 ${typeof v} 转换为数字` }
    }

    case 'boolean': {
      if (typeof v === 'boolean') return { value: v }
      if (typeof v === 'number') return { value: v !== 0 }
      if (typeof v === 'string') {
        const l = v.trim().toLowerCase()
        if (['true', '1', 'yes'].includes(l)) return { value: true }
        if (['false', '0', 'no'].includes(l)) return { value: false }
        return { value: v, error: `"${v}" 无法转换为布尔值` }
      }
      return { value: v, error: `无法将 ${typeof v} 转换为布尔值` }
    }

    case 'date': {
      if (v instanceof Date) {
        return isNaN(v.getTime()) ? { value: v, error: '无效日期' } : { value: formatDate(v) }
      }
      if (typeof v === 'number') {
        const d = v > 100000 ? new Date(v) : excelSerialDateToJSDate(v)
        return isNaN(d.getTime()) ? { value: v, error: `无效日期数值: ${v}` } : { value: formatDate(d) }
      }
      if (typeof v === 'string') {
        if (!v.trim()) return { value: null }
        const d = new Date(v.trim())
        return isNaN(d.getTime()) ? { value: v, error: `"${v}" 无法解析为日期` } : { value: formatDate(d) }
      }
      return { value: v, error: `无法将 ${typeof v} 转换为日期` }
    }
  }
}
