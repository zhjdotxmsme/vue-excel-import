import { describe, it, expect } from 'vitest'
import { flattenColumns, setNested, getNested } from './column'
import type { ColumnConfig } from '../types'

describe('flattenColumns', () => {
  it('returns leaf columns from a flat list', () => {
    const cols: ColumnConfig[] = [
      { label: '姓名', field: 'name', type: 'string' },
      { label: '年龄', field: 'age', type: 'number' },
    ]
    expect(flattenColumns(cols)).toHaveLength(2)
  })

  it('extracts leaves from a nested tree', () => {
    const cols: ColumnConfig[] = [
      {
        label: '基本信息',
        children: [
          { label: '姓名', field: 'basic.name', type: 'string' },
          { label: '年龄', field: 'basic.age', type: 'number' },
        ]
      },
      {
        label: '联系方式',
        children: [
          { label: '邮箱', field: 'contact.email', type: 'string' },
        ]
      }
    ]
    const leaves = flattenColumns(cols)
    expect(leaves).toHaveLength(3)
    expect(leaves.map(l => l.field)).toEqual(['basic.name', 'basic.age', 'contact.email'])
  })

  it('skips group nodes without field', () => {
    const cols: ColumnConfig[] = [
      { label: '分组', children: [
        { label: '子项', field: 'child', type: 'string' }
      ]}
    ]
    const leaves = flattenColumns(cols)
    expect(leaves).toHaveLength(1)
    expect(leaves[0].field).toBe('child')
  })

  it('returns empty array for empty input', () => {
    expect(flattenColumns([])).toEqual([])
  })
})

describe('setNested', () => {
  it('sets a flat key', () => {
    const obj: any = {}
    setNested(obj, 'name', '张三')
    expect(obj.name).toBe('张三')
  })

  it('creates nested objects for dot-paths', () => {
    const obj: any = {}
    setNested(obj, 'basic.name', '张三')
    expect(obj.basic.name).toBe('张三')
  })

  it('sets multiple nested paths', () => {
    const obj: any = {}
    setNested(obj, 'basic.name', '张三')
    setNested(obj, 'basic.age', 25)
    setNested(obj, 'contact.email', 'z@x.com')
    expect(obj).toEqual({
      basic: { name: '张三', age: 25 },
      contact: { email: 'z@x.com' }
    })
  })

  it('overwrites existing nested values', () => {
    const obj: any = { basic: { name: '旧' } }
    setNested(obj, 'basic.name', '新')
    expect(obj.basic.name).toBe('新')
  })
})

describe('getNested', () => {
  it('gets a flat key', () => {
    expect(getNested({ name: '张三' }, 'name')).toBe('张三')
  })

  it('gets a nested path', () => {
    expect(getNested({ basic: { name: '张三' } }, 'basic.name')).toBe('张三')
  })

  it('returns undefined for missing path', () => {
    expect(getNested({}, 'missing.key')).toBeUndefined()
  })

  it('returns undefined for null intermediate', () => {
    expect(getNested({ basic: null }, 'basic.name')).toBeUndefined()
  })
})