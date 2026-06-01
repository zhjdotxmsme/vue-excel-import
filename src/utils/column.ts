import type { ColumnConfig } from '../types'

/** Flatten a tree ColumnConfig into a list of leaf ColumnConfigs only */
export function flattenColumns(columns: ColumnConfig[]): ColumnConfig[] {
  const result: ColumnConfig[] = []
  function walk(list: ColumnConfig[]) {
    for (const col of list) {
      if (col.children && col.children.length > 0) {
        walk(col.children)
      } else if (col.field) {
        result.push(col)
      }
    }
  }
  walk(columns)
  return result
}

/** Get a nested value using dot-path, e.g. getNested(obj, 'basic.name') */
export function getNested(obj: Record<string, any>, path: string): any {
  return path.split('.').reduce((acc, key) => (acc != null ? acc[key] : undefined), obj)
}

/** Set a nested value using dot-path, e.g. setNested(obj, 'basic.name', '张三') */
export function setNested(obj: Record<string, any>, path: string, value: any): void {
  const keys = path.split('.')
  let current = obj
  for (let i = 0; i < keys.length - 1; i++) {
    if (!current[keys[i]] || typeof current[keys[i]] !== 'object') {
      current[keys[i]] = {}
    }
    current = current[keys[i]]
  }
  current[keys[keys.length - 1]] = value
}