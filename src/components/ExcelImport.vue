<template>
  <div class="excel-import">
    <ExcelUploader
      :show-template-button="showTemplateButton"
      @file-selected="onFileSelected"
      @file-cleared="onFileCleared"
      @download-template="onDownloadTemplate"
    />

    <ExcelPreview
      v-if="validationResult.rows.length > 0"
      :columns="columns"
      :rows="validationResult.rows"
      :parse-errors="parseErrors"
      @edit-cell="openCellEdit"
    />

    <ExcelCellEdit
      :visible="editDialogVisible"
      :row-index="editRowIndex"
      :field="editField"
      :value="editValue"
      :current-errors="editErrors"
      :row-data="editRowData"
      @save="onCellSave"
      @close="editDialogVisible = false"
    />

    <SubmitBar
      :total="validationResult.total"
      :valid-count="validationResult.validCount"
      :error-count="validationResult.errorCount"
      :submitting="submitting"
      :progress="submitProgress"
      @submit="onSubmit"
    />

    <el-alert
      v-if="submitError"
      :title="submitError"
      type="error"
      show-icon
      closable
      @close="submitError = ''"
      style="margin-top: 12px"
    />

    <el-alert
      v-if="successMessage"
      :title="successMessage"
      type="success"
      show-icon
      closable
      @close="successMessage = ''"
      style="margin-top: 12px"
    />
  </div>
</template>

<script setup lang="ts">
import { ref, reactive } from 'vue'
import { ElMessage } from 'element-plus'
import type {
  ColumnConfig, ParseResult, ValidationResult,
  RowWithErrors, CellError, SubmitProgress, SubmitResult
} from '../types'
import { useExcelParser } from '../composables/useExcelParser'
import { useExcelValidator } from '../composables/useExcelValidator'
import { useExcelSubmitter } from '../composables/useExcelSubmitter'
import { useExcelTemplate, downloadTemplateBlob } from '../composables/useExcelTemplate'
import ExcelUploader from './ExcelUploader.vue'
import ExcelPreview from './ExcelPreview.vue'
import ExcelCellEdit from './ExcelCellEdit.vue'
import SubmitBar from './SubmitBar.vue'

const props = withDefaults(defineProps<{
  columns: ColumnConfig[]
  submitApi: string | ((data: Record<string, any>[]) => Promise<any>)
  rollbackApi?: string | ((batchId: string) => Promise<any>)
  batchSize?: number
  maxFileSize?: number
  showTemplateButton?: boolean
  templateFileName?: string
  templateSheetName?: string
  templateDescription?: string
  templateExampleRows?: number
}>(), {
  batchSize: 500,
  maxFileSize: 50,
  showTemplateButton: true,
  templateFileName: '导入模板.xlsx',
  templateSheetName: 'Sheet1',
  templateExampleRows: 0
})

const emit = defineEmits<{
  (e: 'parsed', result: ParseResult): void
  (e: 'progress', progress: SubmitProgress): void
  (e: 'success', result: SubmitResult): void
  (e: 'error', error: Error): void
}>()

// State
const submitting = ref(false)
const parseErrors = ref<CellError[]>([])
const validationResult = reactive<ValidationResult>({
  valid: true, total: 0, validCount: 0, errorCount: 0, errors: [], rows: []
})
const submitProgress = reactive<SubmitProgress>({ current: 0, total: 0, batch: 0, totalBatches: 0 })
const submitError = ref('')
const successMessage = ref('')

// Cell edit state
const editDialogVisible = ref(false)
const editRowIndex = ref(0)
const editField = ref('')
const editValue = ref('')
const editErrors = ref<CellError[]>([])
const editRowData = ref<Record<string, any> | null>(null)

let currentFile: File | null = null
let currentRows: Record<string, any>[] = []

async function onFileSelected(file: File) {
  const maxBytes = props.maxFileSize! * 1024 * 1024
  if (file.size > maxBytes) {
    ElMessage.warning(`文件大小不能超过 ${props.maxFileSize}MB`)
    return
  }

  currentFile = file

  try {
    const parseResult = await useExcelParser(file, props.columns)
    currentRows = parseResult.rows
    parseErrors.value = parseResult.parseErrors

    const vResult = useExcelValidator(parseResult.rows, props.columns)
    Object.assign(validationResult, vResult)

    emit('parsed', parseResult)
  } catch (err: any) {
    ElMessage.error(`解析失败: ${err.message}`)
  }
}

function onFileCleared() {
  currentFile = null
  currentRows = []
  parseErrors.value = []
  validationResult.total = 0
  validationResult.validCount = 0
  validationResult.errorCount = 0
  validationResult.valid = true
  validationResult.errors = []
  validationResult.rows = []
  successMessage.value = ''
  submitError.value = ''
}

async function onDownloadTemplate() {
  try {
    const blob = await useExcelTemplate(props.columns, {
      sheetName: props.templateSheetName,
      description: props.templateDescription,
      exampleRows: props.templateExampleRows
    })
    downloadTemplateBlob(blob, props.templateFileName)
  } catch (err: any) {
    ElMessage.error(`模板生成失败: ${err.message}`)
  }
}

function openCellEdit(rowIndex: number, field: string) {
  const row = validationResult.rows[rowIndex]
  if (!row) return
  editRowIndex.value = rowIndex
  editField.value = field
  editValue.value = row.data[field] ?? ''
  editErrors.value = row.errors.filter(e => e.field === field)
  editRowData.value = row.data
  editDialogVisible.value = true
}

function onCellSave(rowIndex: number, field: string, value: any) {
  const row = validationResult.rows[rowIndex]
  if (!row) return
  row.data[field] = value

  // Re-validate this single row
  const singleResult = useExcelValidator([row.data], props.columns)
  row.errors = singleResult.rows[0]?.errors ?? []
  row.valid = row.errors.length === 0

  // Recalculate totals
  const validCount = validationResult.rows.filter(r => r.valid).length
  validationResult.validCount = validCount
  validationResult.errorCount = validationResult.total - validCount
  validationResult.valid = validCount === validationResult.total
  validationResult.errors = validationResult.rows.flatMap(r => r.errors)

  editDialogVisible.value = false
}

async function onSubmit() {
  const validRows = validationResult.rows
    .filter(r => r.valid)
    .map(r => r.data)

  if (validRows.length === 0) return

  submitting.value = true
  submitError.value = ''
  successMessage.value = ''

  try {
    const result = await useExcelSubmitter(validRows, {
      submitApi: props.submitApi,
      rollbackApi: props.rollbackApi,
      batchSize: props.batchSize,
      onProgress: (p) => {
        Object.assign(submitProgress, p)
        emit('progress', p)
      }
    })

    if (result.success) {
      successMessage.value = '数据导入成功！'
      emit('success', result)
    } else {
      const msg = result.rollbackSuccess
        ? `第 ${result.failedBatch} 批提交失败，已回滚`
        : `第 ${result.failedBatch} 批提交失败，回滚也失败: ${result.error}`
      submitError.value = msg
      emit('error', new Error(msg))
    }
  } catch (err: any) {
    submitError.value = `提交异常: ${err.message}`
    emit('error', err)
  } finally {
    submitting.value = false
  }
}
</script>

<style scoped>
.excel-import {
  max-width: 960px;
  margin: 0 auto;
}
</style>
