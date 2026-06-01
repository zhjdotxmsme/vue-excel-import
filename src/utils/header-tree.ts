import type { ColumnConfig, CellValue } from '../types'

export interface MergeRange {
  startRow: number
  startCol: number
  endRow: number
  endCol: number
}

export interface HeaderTreeOptions {
  merges?: MergeRange[]
}

interface ColPath {
  labels: string[]
  colIdx: number
}

/**
 * Detect multi-level header tree from Excel sheet rows + merges.
 *
 * Algorithm:
 * 1. Determine header row count by scanning for the last row that contains
 *    at least one "active" header cell (non-empty, and either top-left of a
 *    merge, or not inside any merge).
 * 2. If single header row → return flat list of labels (backward compatible).
 * 3. Build merge lookup map: "row,col" → MergeRange.
 * 4. For each column index collect a label path from top to bottom.
 *    - Empty cells that sit inside a merged region (but are not the top-left
 *      cell) inherit the merge's label, so sub-columns carry their parent's
 *      group name.
 *    - Non‑empty cells always use their own value.
 *    - Consecutive duplicates are collapsed.
 * 5. Aggregate column paths into a tree via buildTreeFromPaths.
 */
export function detectHeaderTree(
  rows: CellValue[][],
  options?: HeaderTreeOptions
): ColumnConfig[] {
  if (!rows || rows.length === 0) return []

  const merges = options?.merges ?? []

  // 3. Build merge lookup map early (needed for header-row detection)
  const mergeMap = new Map<string, MergeRange>()
  for (const m of merges) {
    for (let r = m.startRow; r <= m.endRow; r++) {
      for (let c = m.startCol; c <= m.endCol; c++) {
        mergeMap.set(`${r},${c}`, m)
      }
    }
  }

  // 1. Determine header row count
  let headerRowCount = 1
  if (merges.length > 0) {
    // Find the last row that has an "active" header cell:
    // a non‑empty cell that is either top‑left of a merge or not inside any merge.
    let lastRowWithHeader = 0
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r]
      if (!row) continue
      let hasHeader = false
      for (let c = 0; c < row.length; c++) {
        const cell = row[c]
        if (cell == null || String(cell).trim() === '') continue
        const mr = mergeMap.get(`${r},${c}`)
        // Skip if it's inside a merge but not top-left
        if (mr && (r !== mr.startRow || c !== mr.startCol)) continue
        hasHeader = true
        break
      }
      if (hasHeader) lastRowWithHeader = r
    }
    headerRowCount = lastRowWithHeader + 1
  }
  headerRowCount = Math.min(headerRowCount, rows.length)

  // 2. Single header row — flat list
  if (headerRowCount === 1) {
    const headerRow = rows[0]
    if (!headerRow) return []
    const result: ColumnConfig[] = []
    headerRow.forEach((cell, colIdx) => {
      const label = String(cell ?? '').trim()
      if (label) {
        result.push({ label, _colIdx: colIdx } as ColumnConfig)
      }
    })
    return result
  }

  // 4. Determine max column index in the header area
  let maxCol = 0
  for (let r = 0; r < headerRowCount; r++) {
    const row = rows[r]
    if (row && row.length > maxCol) maxCol = row.length
  }

  // 5. Build column paths with merge-label inheritance
  const colPaths: ColPath[] = []

  for (let colIdx = 0; colIdx < maxCol; colIdx++) {
    const labels: string[] = []

    for (let rowIdx = 0; rowIdx < headerRowCount; rowIdx++) {
      const cell = rows[rowIdx]?.[colIdx]
      const merge = mergeMap.get(`${rowIdx},${colIdx}`)

      let addLabel: string | null = null

      if (merge) {
        if (rowIdx === merge.startRow && colIdx === merge.startCol) {
          // ── Top-left of a merged cell ──
          const val = String(cell ?? '').trim()
          if (val) addLabel = val
        } else {
          // ── Inside a merged cell, not the top-left ──
          const val = String(cell ?? '').trim()
          if (val) {
            // Cell has its own value (e.g. a sub-header physically inside the
            // merge area)
            addLabel = val
          } else {
            // Inherit the merge's label
            const topLeftCell = rows[merge.startRow]?.[merge.startCol]
            const inherited = String(topLeftCell ?? '').trim()
            // Prevent consecutive duplicates
            if (inherited && inherited !== labels[labels.length - 1]) {
              addLabel = inherited
            }
          }
        }
      } else {
        // ── Not part of any merged cell ──
        const val = String(cell ?? '').trim()
        if (val) addLabel = val
      }

      if (addLabel !== null) {
        labels.push(addLabel)
      }
    }

    if (labels.length > 0) {
      colPaths.push({ labels, colIdx })
    }
  }

  // 6. Build tree from paths
  return buildTreeFromPaths(colPaths)
}

// ── Internal helpers ────────────────────────────────────────────────

/**
 * Build a tree of ColumnConfig from column label paths.
 *
 * Each path represents one data column's header labels from top-level group
 * down to the leaf name.  Paths that share a common prefix are nested under
 * the same group node.
 */
function buildTreeFromPaths(paths: ColPath[]): ColumnConfig[] {
  if (paths.length === 0) return []

  // Group by root label
  const rootMap = new Map<string, ColPath[]>()
  for (const p of paths) {
    const root = p.labels[0]
    if (!rootMap.has(root)) rootMap.set(root, [])
    rootMap.get(root)!.push(p)
  }

  const result: ColumnConfig[] = []

  for (const [rootLabel, entries] of rootMap) {
    if (entries.length === 1 && entries[0].labels.length === 1) {
      // Single leaf column
      result.push({ label: rootLabel, _colIdx: entries[0].colIdx } as ColumnConfig)
    } else if (entries.every(e => e.labels.length === 1)) {
      // Multiple leaves all at depth 1 — flat list
      for (const e of entries) {
        result.push({ label: e.labels[0], _colIdx: e.colIdx } as ColumnConfig)
      }
    } else {
      // Has nested children
      const children: ColumnConfig[] = []
      const deepEntries: ColPath[] = []

      for (const e of entries) {
        if (e.labels.length === 1) {
          children.push({ label: e.labels[0], _colIdx: e.colIdx } as ColumnConfig)
        } else if (e.labels.length === 2) {
          children.push({ label: e.labels[1], _colIdx: e.colIdx } as ColumnConfig)
        } else {
          // Depth ≥ 3 — collect for grouped recursion so that entries
          // sharing the same second-level label get merged into one group
          deepEntries.push({
            labels: e.labels.slice(1),
            colIdx: e.colIdx
          })
        }
      }

      if (deepEntries.length > 0) {
        children.push(...buildTreeFromPaths(deepEntries))
      }

      result.push({ label: rootLabel, children } as ColumnConfig)
    }
  }

  return result
}

// ── Column matching ────────────────────────────────────────────────

export interface MatchResult {
  /** Map of Excel column index → ColumnConfig leaf (columnIndex 0-based) */
  columnMap: Map<number, ColumnConfig>
  /** Labels in ColumnConfig that were NOT found in the Excel header tree */
  missingColumns: string[]
  /** Labels in Excel that were NOT matched to any ColumnConfig */
  unmatchedHeaders: string[]
}

/**
 * Collect all leaf labels from a ColumnConfig tree.
 * A leaf is a node that has `field` defined (i.e. a data column).
 */
export function collectLeafLabels(nodes: ColumnConfig[]): string[] {
  const result: string[] = []
  for (const node of nodes) {
    if (node.children && node.children.length > 0) {
      result.push(...collectLeafLabels(node.children))
    } else if (node.field) {
      result.push(node.label)
    }
  }
  return result
}

/**
 * Match a user's ColumnConfig against a detected Excel header tree.
 *
 * Traverses both trees in parallel by `label`.  For each Excel leaf node
 * that finds a matching user-config leaf, the column index (`_colIdx`) is
 * recorded in the returned `columnMap`.
 */
export function matchColumnTree(
  excelTree: ColumnConfig[],
  userConfig: ColumnConfig[]
): MatchResult {
  const columnMap = new Map<number, ColumnConfig>()
  const unmatchedHeaders: string[] = []
  const matchedConfigLabels = new Set<string>()

  function walk(excelNodes: ColumnConfig[], configNodes: ColumnConfig[]): void {
    for (const excelNode of excelNodes) {
      const configNode = configNodes.find(c => c.label === excelNode.label)

      if (!configNode) {
        // No matching config node
        if (!excelNode.children || excelNode.children.length === 0) {
          // Leaf excel node → unmatched header
          unmatchedHeaders.push(excelNode.label)
        }
        // Group node with no config match: skip entirely (no recursion)
        continue
      }

      if (configNode.children && configNode.children.length > 0) {
        // Config node is a group — recurse into children
        walk(excelNode.children || [], configNode.children)
      } else if (configNode.field) {
        // Config node is a leaf — extract column index
        const colIdx = (excelNode as any)._colIdx
        if (colIdx !== undefined) {
          columnMap.set(colIdx, configNode)
        }
        matchedConfigLabels.add(configNode.label)
      }
    }
  }

  walk(excelTree, userConfig)

  // Determine missing columns: config leaf labels never matched
  const allConfigLabels = collectLeafLabels(userConfig)
  const missingColumns = allConfigLabels.filter(l => !matchedConfigLabels.has(l))

  return { columnMap, missingColumns, unmatchedHeaders }
}
