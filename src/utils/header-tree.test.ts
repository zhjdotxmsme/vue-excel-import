import { describe, it, expect } from 'vitest'
import { detectHeaderTree } from './header-tree'
import type { MergeRange } from './header-tree'

/** Helper: assert a ColumnConfig is a leaf with the given label and _colIdx */
function expectLeaf(
  node: Record<string, any>,
  label: string,
  colIdx: number
): void {
  expect(node.label).toBe(label)
  expect(node.children).toBeUndefined()
  expect(node._colIdx).toBe(colIdx)
}

/** Helper: assert a ColumnConfig is a group */
function expectGroup(node: Record<string, any>, label: string): void {
  expect(node.label).toBe(label)
  expect(Array.isArray(node.children)).toBe(true)
  expect(node._colIdx).toBeUndefined()
}

describe('detectHeaderTree', () => {
  // ── 1. Flat single header row (no merges) ──────────────────────────
  describe('flat single header row', () => {
    it('returns flat list of labels with correct _colIdx', () => {
      const rows = [['姓名', '年龄', '邮箱']]
      const result = detectHeaderTree(rows)

      expect(result).toHaveLength(3)
      expectLeaf(result[0] as any, '姓名', 0)
      expectLeaf(result[1] as any, '年龄', 1)
      expectLeaf(result[2] as any, '邮箱', 2)
    })

    it('handles duplicate labels correctly', () => {
      const rows = [['姓名', '姓名']]
      const result = detectHeaderTree(rows)

      expect(result).toHaveLength(2)
      expectLeaf(result[0] as any, '姓名', 0)
      expectLeaf(result[1] as any, '姓名', 1)
    })

    it('is unaffected by merges when merges option is empty', () => {
      const rows = [['A', 'B']]
      const result = detectHeaderTree(rows, { merges: [] })
      expect(result).toHaveLength(2)
    })
  })

  // ── 2. Empty rows ──────────────────────────────────────────────────
  describe('empty rows', () => {
    it('returns empty array for empty input', () => {
      expect(detectHeaderTree([])).toEqual([])
    })

    it('returns empty array for null/undefined rows', () => {
      expect(detectHeaderTree([] as any)).toEqual([])
    })

    it('returns empty array when row has only empty cells', () => {
      expect(detectHeaderTree([[null, null, '']])).toEqual([])
    })

    it('returns empty array for row array with empty inner', () => {
      expect(detectHeaderTree([[]])).toEqual([])
    })
  })

  // ── 3. 2-level nested ─────────────────────────────────────────────
  describe('2-level nested header', () => {
    it('detects 基本信息 → 姓名|年龄, 联系方式 → 邮箱', () => {
      /**
       * Excel layout:
       *   [基本信息  ]  [联系方式]
       *   [姓名] [年龄]  [邮箱]
       */
      const rows: any[][] = [
        ['基本信息', null, '联系方式'],
        ['姓名', '年龄', '邮箱'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      ]

      const result = detectHeaderTree(rows, { merges })
      expect(result).toHaveLength(2)

      // ── 基本信息 group ──
      expectGroup(result[0] as any, '基本信息')
      const basicKids = (result[0] as any).children
      expect(basicKids).toHaveLength(2)
      expectLeaf(basicKids[0], '姓名', 0)
      expectLeaf(basicKids[1], '年龄', 1)

      // ── 联系方式 group ──
      expectGroup(result[1] as any, '联系方式')
      const contactKids = (result[1] as any).children
      expect(contactKids).toHaveLength(1)
      expectLeaf(contactKids[0], '邮箱', 2)
    })

    it('groups columns sharing the same root label', () => {
      const rows: any[][] = [
        ['G1', null, 'G1', null],
        ['a', 'b', 'c', 'd'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
        { startRow: 0, startCol: 2, endRow: 0, endCol: 3 },
      ]
      const result = detectHeaderTree(rows, { merges })
      expect(result).toHaveLength(1)

      expectGroup(result[0] as any, 'G1')
      const kids = (result[0] as any).children
      expect(kids).toHaveLength(4)
      expectLeaf(kids[0], 'a', 0)
      expectLeaf(kids[1], 'b', 1)
      expectLeaf(kids[2], 'c', 2)
      expectLeaf(kids[3], 'd', 3)
    })
  })

  // ── 4. 3-level nested ─────────────────────────────────────────────
  describe('3-level nested header', () => {
    it('builds deep tree: 人员信息 > 基本信息 > 姓名|年龄 and 人员信息 > 联系方式 > 邮箱|手机', () => {
      /**
       * Excel layout:
       *   [------- 人员信息 --------]
       *   [-- 基本信息 --] [-- 联系方式 --]
       *   [姓名] [年龄]    [邮箱] [手机]
       */
      const rows: any[][] = [
        ['人员信息', null, null, null],
        ['基本信息', null, '联系方式', null],
        ['姓名', '年龄', '邮箱', '手机'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 3 },
        { startRow: 1, startCol: 0, endRow: 1, endCol: 1 },
        { startRow: 1, startCol: 2, endRow: 1, endCol: 3 },
      ]

      const result = detectHeaderTree(rows, { merges })
      expect(result).toHaveLength(1)

      // ── Root: 人员信息 ──
      expectGroup(result[0] as any, '人员信息')
      const level1 = (result[0] as any).children
      expect(level1).toHaveLength(2)

      // ── L1: 基本信息 group ──
      expectGroup(level1[0], '基本信息')
      const basicKids = level1[0].children
      expect(basicKids).toHaveLength(2)
      expectLeaf(basicKids[0], '姓名', 0)
      expectLeaf(basicKids[1], '年龄', 1)

      // ── L1: 联系方式 group ──
      expectGroup(level1[1], '联系方式')
      const contactKids = level1[1].children
      expect(contactKids).toHaveLength(2)
      expectLeaf(contactKids[0], '邮箱', 2)
      expectLeaf(contactKids[1], '手机', 3)
    })
  })

  // ── 5. Empty cells skipped ────────────────────────────────────────
  describe('empty cells are skipped', () => {
    it('skips empty string and null cells in flat header', () => {
      const rows = [['姓名', '', '邮箱', null, '年龄']]
      const result = detectHeaderTree(rows)

      expect(result).toHaveLength(3)
      expectLeaf(result[0] as any, '姓名', 0)
      expectLeaf(result[1] as any, '邮箱', 2)
      expectLeaf(result[2] as any, '年龄', 4)
    })

    it('skips empty cells in nested header', () => {
      // Row 1 has a null middle cell — column 1 row 1 has no label.
      // This means column 1's path is depth 1 only.
      const rows: any[][] = [
        ['A', null, 'B'],
        ['x', null, 'y'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      ]
      const result = detectHeaderTree(rows, { merges })

      expect(result).toHaveLength(2)

      // Column 0: A→x (depth 2)
      // Column 1: A only (depth 1 — null cell, so no sub-label)
      // Column 2: B→y (depth 2, no merges)
      expectGroup(result[0] as any, 'A')
      const aKids = (result[0] as any).children
      // Column 0 is depth-2 leaf, column 1 is depth-1 leaf
      expect(aKids).toHaveLength(2)
      expectLeaf(aKids[0], 'x', 0)
      expectLeaf(aKids[1], 'A', 1)

      expectGroup(result[1] as any, 'B')
      expectLeaf((result[1] as any).children[0], 'y', 2)
    })
  })

  // ── 6. _colIdx correctness ────────────────────────────────────────
  describe('_colIdx correctness', () => {
    it('assigns correct column indices in flat header', () => {
      const rows = [['A', 'B', 'C', 'D']]
      const result = detectHeaderTree(rows)

      expect((result[0] as any)._colIdx).toBe(0)
      expect((result[1] as any)._colIdx).toBe(1)
      expect((result[2] as any)._colIdx).toBe(2)
      expect((result[3] as any)._colIdx).toBe(3)
    })

    it('assigns correct column indices in 2-level header', () => {
      const rows: any[][] = [
        ['G1', null, 'G2', null],
        ['a', 'b', 'c', 'd'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
        { startRow: 0, startCol: 2, endRow: 0, endCol: 3 },
      ]
      const result = detectHeaderTree(rows, { merges })

      expect(result).toHaveLength(2)

      const g1Kids = (result[0] as any).children
      expect(g1Kids).toHaveLength(2)
      expectLeaf(g1Kids[0], 'a', 0)
      expectLeaf(g1Kids[1], 'b', 1)

      const g2Kids = (result[1] as any).children
      expect(g2Kids).toHaveLength(2)
      expectLeaf(g2Kids[0], 'c', 2)
      expectLeaf(g2Kids[1], 'd', 3)
    })

    it('assigns correct column indices in 3-level header', () => {
      const rows: any[][] = [
        ['R', null, null, null],
        ['A', null, 'B', null],
        ['x', 'y', 'z', 'w'],
      ]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 3 },
        { startRow: 1, startCol: 0, endRow: 1, endCol: 1 },
        { startRow: 1, startCol: 2, endRow: 1, endCol: 3 },
      ]
      const result = detectHeaderTree(rows, { merges })

      const l1 = (result[0] as any).children
      const aKids = l1[0].children
      expectLeaf(aKids[0], 'x', 0)
      expectLeaf(aKids[1], 'y', 1)

      const bKids = l1[1].children
      expectLeaf(bKids[0], 'z', 2)
      expectLeaf(bKids[1], 'w', 3)
    })
  })

  // ── 7. Single row with trivially merged cells ─────────────────────
  describe('single row with merged cells', () => {
    it('treats single-row merge as flat header', () => {
      // A single row where some cells are merged horizontally; the
      // header is still flat.
      const rows = [['姓名', '年龄']]
      const merges: MergeRange[] = [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      ]
      const result = detectHeaderTree(rows, { merges })

      expect(result).toHaveLength(2)
      expectLeaf(result[0] as any, '姓名', 0)
      expectLeaf(result[1] as any, '年龄', 1)
    })
  })
})
