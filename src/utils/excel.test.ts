import { describe, it, expect } from 'vitest'
import { excelSerialDateToJSDate, formatDate, convertValue } from './excel'

describe('excelSerialDateToJSDate', () => {
  it('converts Excel serial 45291 to 2024-01-01', () => {
    const d = excelSerialDateToJSDate(45291)
    expect(d.getFullYear()).toBe(2024)
    expect(d.getMonth()).toBe(0)
    expect(d.getDate()).toBe(1)
  })

  it('converts Excel serial 1 to 1900-01-01', () => {
    const d = excelSerialDateToJSDate(1)
    expect(d.getFullYear()).toBe(1900)
    expect(d.getMonth()).toBe(0)
    expect(d.getDate()).toBe(1)
  })
})

describe('formatDate', () => {
  it('formats Date to YYYY-MM-DD', () => {
    expect(formatDate(new Date(2024, 0, 15))).toBe('2024-01-15')
  })

  it('pads month and day with zeros', () => {
    expect(formatDate(new Date(2024, 2, 5))).toBe('2024-03-05')
  })
})

describe('convertValue', () => {
  describe('type=string', () => {
    it('trims whitespace', () => {
      expect(convertValue('  hi  ', 'string').value).toBe('hi')
    })
    it('converts number to string', () => {
      expect(convertValue(25, 'string').value).toBe('25')
    })
    it('converts boolean to string', () => {
      expect(convertValue(true, 'string').value).toBe('true')
    })
    it('converts Date to YYYY-MM-DD', () => {
      expect(convertValue(new Date(2024, 0, 15), 'string').value).toBe('2024-01-15')
    })
    it('returns empty string for null', () => {
      expect(convertValue(null, 'string').value).toBe('')
    })
  })

  describe('type=number', () => {
    it('keeps number as-is', () => {
      expect(convertValue(42, 'number').value).toBe(42)
    })
    it('converts boolean true to 1', () => {
      expect(convertValue(true, 'number').value).toBe(1)
    })
    it('converts boolean false to 0', () => {
      expect(convertValue(false, 'number').value).toBe(0)
    })
    it('parses numeric string', () => {
      expect(convertValue('42', 'number').value).toBe(42)
    })
    it('errors on non-numeric string', () => {
      expect(convertValue('abc', 'number').error).toBeTruthy()
    })
  })

  describe('type=boolean', () => {
    it('keeps boolean as-is', () => {
      expect(convertValue(true, 'boolean').value).toBe(true)
    })
    it('converts 0 to false, non-0 to true', () => {
      expect(convertValue(0, 'boolean').value).toBe(false)
      expect(convertValue(1, 'boolean').value).toBe(true)
    })
    it('converts "true"/"1"/"yes" to true', () => {
      expect(convertValue('true', 'boolean').value).toBe(true)
      expect(convertValue('1', 'boolean').value).toBe(true)
      expect(convertValue('yes', 'boolean').value).toBe(true)
    })
    it('converts "false"/"0"/"no" to false', () => {
      expect(convertValue('false', 'boolean').value).toBe(false)
      expect(convertValue('0', 'boolean').value).toBe(false)
      expect(convertValue('no', 'boolean').value).toBe(false)
    })
    it('errors on ambiguous string', () => {
      expect(convertValue('maybe', 'boolean').error).toBeTruthy()
    })
  })

  describe('type=date', () => {
    it('formats Date object', () => {
      expect(convertValue(new Date(2024, 5, 15), 'date').value).toBe('2024-06-15')
    })
    it('handles Unix timestamp (number > 100000)', () => {
      // 2024-01-01T00:00:00.000Z in ms
      expect(convertValue(1704067200000, 'date').value).toBe('2024-01-01')
    })
    it('handles Excel serial number (number <= 100000)', () => {
      expect(convertValue(45291, 'date').value).toBe('2024-01-01')
    })
    it('parses date string', () => {
      expect(convertValue('2024-06-15', 'date').value).toBe('2024-06-15')
    })
    it('errors on invalid string', () => {
      expect(convertValue('not-a-date', 'date').error).toBeTruthy()
    })
  })
})
