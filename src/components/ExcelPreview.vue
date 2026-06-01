<template>
  <div class="excel-preview" v-if="rows.length > 0">
    <div class="preview-header">
      <span class="preview-title">数据预览</span>
      <span class="preview-count">共 {{ total }} 条记录</span>
    </div>

    <el-table
      :data="paginatedRows"
      border
      stripe
      max-height="480"
      :cell-style="cellStyle"
      size="small"
      highlight-current-row
    >
      <el-table-column type="index" label="#" width="50" fixed />
      <el-table-column
        v-for="col in visibleColumns"
        :key="col.field"
        :prop="col.field"
        :label="col.label"
        :width="colWidth(col)"
        show-overflow-tooltip
      >
        <template #default="{ row: r, $index }">
          <div
            class="cell-wrapper"
            :class="{ 'cell-error': hasCellError(r, col.field), 'cell-required': col.required }"
            @dblclick="startEdit(r, col.field, $index)"
          >
            <span>{{ getNested(r.data, col.field) ?? '' }}</span>
            <el-tooltip
              v-if="hasCellError(r, col.field)"
              :content="getCellErrors(r, col.field).map((e: any) => e.message).join('; ')"
              placement="top"
              effect="dark"
            >
              <span class="error-badge">!</span>
            </el-tooltip>
          </div>
        </template>
      </el-table-column>

      <el-table-column label="校验状态" width="130" fixed="right">
        <template #default="{ row: r }">
          <el-tag v-if="r.valid" type="success" size="small">通过</el-tag>
          <el-tag v-else type="danger" size="small">
            {{ r.errors.length }} 个错误
          </el-tag>
        </template>
      </el-table-column>
    </el-table>

    <div class="pagination-wrapper" v-if="rows.length > pageSize">
      <el-pagination
        v-model:current-page="currentPage"
        :page-size="pageSize"
        :total="rows.length"
        layout="prev, pager, next"
        small
      />
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import type { ColumnConfig, RowWithErrors } from '../types'
import { getNested, flattenColumns } from '../utils/column'

const props = defineProps<{
  columns: ColumnConfig[]
  rows: RowWithErrors[]
  parseErrors: any[]
}>()

const emit = defineEmits<{
  (e: 'edit-cell', rowIndex: number, field: string): void
}>()

const currentPage = ref(1)
const pageSize = 100

const visibleColumns = computed(() => flattenColumns(props.columns))
const total = computed(() => props.rows.length)

const paginatedRows = computed(() => {
  const start = (currentPage.value - 1) * pageSize
  return props.rows.slice(start, start + pageSize)
})

function colWidth(col: ColumnConfig): number | undefined {
  switch (col.type) {
    case 'number': return 100
    case 'date': return 120
    case 'boolean': return 90
    default: return undefined
  }
}

function hasCellError(row: RowWithErrors, field: string): boolean {
  return row.errors.some(e => e.field === field)
}

function getCellErrors(row: RowWithErrors, field: string) {
  return row.errors.filter(e => e.field === field)
}

function cellStyle({ row, column }: any) {
  const field = column.property
  if (!field) return {}
  const r = row as RowWithErrors
  if (r.errors?.some(e => e.field === field)) {
    return { backgroundColor: '#fef2f2', color: '#dc2626' }
  }
  return {}
}

function startEdit(row: RowWithErrors, field: string, index: number) {
  const globalIdx = (currentPage.value - 1) * pageSize + index
  if (!row.valid) {
    emit('edit-cell', globalIdx, field)
  }
}
</script>

<style scoped>
.excel-preview {
  margin-bottom: 16px;
}

.preview-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}

.preview-title { font-weight: 600; font-size: 15px; }
.preview-count { font-size: 13px; color: #909399; }

.cell-wrapper {
  display: flex;
  align-items: center;
  gap: 4px;
  min-height: 24px;
}

.cell-wrapper.cell-error { cursor: pointer; }
.cell-wrapper.cell-required::after {
  content: '*';
  color: #e74c3c;
  margin-left: 2px;
}

.error-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  background: #dc2626;
  color: white;
  font-size: 10px;
  font-weight: bold;
  cursor: help;
}

.pagination-wrapper {
  display: flex;
  justify-content: center;
  margin-top: 12px;
}
</style>
