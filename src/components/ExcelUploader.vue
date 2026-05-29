<template>
  <div class="excel-uploader">
    <div
      class="upload-area"
      @dragover.prevent="dragging = true"
      @dragleave.prevent="dragging = false"
      @drop.prevent="onDrop"
      @click="triggerInput"
      :class="{ 'is-dragging': dragging }"
    >
      <div v-if="!file" class="upload-placeholder">
        <div class="upload-icon">📁</div>
        <p class="upload-text">拖拽 Excel 文件到此处，或点击上传</p>
        <p class="upload-hint">支持 .xlsx / .xls</p>
        <el-button type="primary" size="large" @click.stop="triggerInput">
          选择文件
        </el-button>
      </div>

      <div v-else class="file-info">
        <div class="file-icon">{{ fileIcon }}</div>
        <div class="file-details">
          <p class="file-name">{{ file.name }}</p>
          <p class="file-size">{{ formatSize(file.size) }}</p>
        </div>
        <el-button size="small" type="danger" plain @click.stop="clearFile">移除</el-button>
      </div>
    </div>

    <input
      ref="fileInput"
      type="file"
      accept=".xlsx,.xls"
      style="display: none"
      @change="onFileChange"
    />

    <div v-if="showTemplateButton" class="template-download">
      <el-button text @click="emit('download-template')">
        📋 下载导入模板
      </el-button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, computed } from 'vue'
import { ElMessage } from 'element-plus'

const props = withDefaults(defineProps<{
  showTemplateButton?: boolean
}>(), {
  showTemplateButton: true
})

const emit = defineEmits<{
  (e: 'file-selected', file: File): void
  (e: 'file-cleared'): void
  (e: 'download-template'): void
}>()

const dragging = ref(false)
const file = ref<File | null>(null)
const fileInput = ref<HTMLInputElement | null>(null)

const fileIcon = computed(() => {
  if (!file.value) return ''
  return file.value.name.endsWith('.xls') ? '📄' : '📊'
})

function triggerInput() {
  fileInput.value?.click()
}

function onFileChange(e: Event) {
  const input = e.target as HTMLInputElement
  if (input.files?.length) {
    selectFile(input.files[0])
  }
}

function onDrop(e: DragEvent) {
  dragging.value = false
  const files = e.dataTransfer?.files
  if (files?.length) {
    selectFile(files[0])
  }
}

function selectFile(f: File) {
  if (!f.name.match(/\.xlsx?$/i)) {
    ElMessage.warning('请选择 .xlsx 或 .xls 文件')
    return
  }
  file.value = f
  emit('file-selected', f)
}

function clearFile() {
  file.value = null
  emit('file-cleared')
  if (fileInput.value) fileInput.value.value = ''
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}
</script>

<style scoped>
.excel-uploader {
  margin-bottom: 16px;
}

.upload-area {
  border: 2px dashed #d9d9d9;
  border-radius: 12px;
  padding: 32px;
  text-align: center;
  cursor: pointer;
  transition: all 0.3s;
}

.upload-area:hover,
.upload-area.is-dragging {
  border-color: #409eff;
  background: rgba(64, 158, 255, 0.04);
}

.upload-icon { font-size: 40px; margin-bottom: 8px; }
.upload-text { font-weight: 600; margin: 4px 0; }
.upload-hint { font-size: 13px; color: #909399; margin: 2px 0 16px; }

.file-info {
  display: flex;
  align-items: center;
  gap: 12px;
}

.file-icon { font-size: 32px; }
.file-details { flex: 1; text-align: left; }
.file-name { font-weight: 600; margin: 0; }
.file-size { font-size: 12px; color: #909399; margin: 2px 0 0; }

.template-download {
  margin-top: 8px;
  text-align: center;
}
</style>
