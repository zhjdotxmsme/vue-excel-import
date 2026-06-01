<template>
  <div class="excel-import">
    <ExcelUploader
      :show-template-button="showTemplateButton"
      :disabled="isParsing"
      @file-selected="onFileSelected"
      @file-cleared="onFileCleared"
      @download-template="onDownloadTemplate"
    />

    <!-- Parsing progress bar -->
    <div v-if="isParsing" class="parsing-progress">
      <p class="parsing-label">正在解析 Excel 文件...</p>
      <el-progress
        :percentage="parsingProgress"
        :stroke-width="20"
        :text-inside="true"
        :status="parsingProgress >= 100 ? 'success' : undefined"
      />
    </div>

    <!-- Sheet info -->
    <div v-if="sheetNames.length > 0" class="sheet-info">
      <span class="sheet-badge">{{ sheetNames.length }} 个 Sheet</span>
      <span class="sheet-detail">{{ sheetNames.join(' · ') }}</span>
    </div>

    <!-- Header-column mismatch warnings -->
    <el-alert
      v-if="missingColumns.length > 0"
      :title="'Excel 中缺少配置列: ' + missingColumns.join('、')"
      type="warning"
      show-icon
      :closable="false"
      description="请检查模板是否正确，缺少的列数据将无法导入"
      style="margin-bottom: 12px"
    />
    <el-alert
      v-if="unmatchedHeaders.length > 0"
      :title="'Excel 中存在未配置的列: ' + unmatchedHeaders.join('、')"
      type="info"
      show-icon
      :closable="false"
      description="这些列将被忽略，如需导入请补充 columns 配置"
      style="margin-bottom: 12px"
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
import { ref, reactive, watch } from 'vue'
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

const parser = useExcelParser()

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
const isParsing = ref(false)
const parseErrors = ref<CellError[]>([])
const parsingProgress = ref(0)
const sheetNames = ref<string[]>([])
const missingColumns = ref<string[]>([])
const unmatchedHeaders = ref<string[]>([])
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

async function onFileSelected(file: File) {
  const maxBytes = props.maxFileSize! * 1024 * 1024
  if (file.size > maxBytes) {
    ElMessage.warning(`文件大小不能超过 ${props.maxFileSize}MB`)
    return
  }

  currentFile = file

  // Reset previous data
  onFileCleared()

  try {
    isParsing.value = true

    // Parse in Web Worker — UI stays responsive
    const parseResult = await parser.parse(file, props.columns)

    parseErrors.value = parseResult.parseErrors
    sheetNames.value = parseResult.sheets ?? []
    missingColumns.value = parseResult.missingColumns ?? []
    unmatchedHeaders.value = parseResult.unmatchedHeaders ?? []

    // Validate
    const vResult = useExcelValidator(parseResult.rows, props.columns)
    Object.assign(validationResult, vResult)

    emit('parsed', parseResult)
  } catch (err: any) {
    ElMessage.error(`解析失败: ${err.message}`)
  } finally {
    isParsing.value = false
  }
}

// Sync progress from parser
watch(parser.progress, (val) => {
  parsingProgress.value = val
})

function onFileCleared() {
  currentFile = null
  parseErrors.value = []
  missingColumns.value = []
  unmatchedHeaders.value = []
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

.parsing-progress {
  margin-bottom: 16px;
  padding: 16px;
  background: #fafafa;
  border-radius: 8px;
}

.parsing-label {
  font-size: 14px;
  color: #606266;
  margin: 0 0 8px 0;
}

.sheet-info {
  margin-bottom: 12px;
  padding: 8px 12px;
  background: #f0f9ff;
  border: 1px solid #bae6fd;
  border-radius: 6px;
  font-size: 13px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.sheet-badge {
  display: inline-flex;
  align-items: center;
  padding: 2px 8px;
  background: #0ea5e9;
  color: white;
  border-radius: 4px;
  font-weight: 600;
  font-size: 12px;
}

.sheet-detail {
  color: #0369a1;
}
</style>
