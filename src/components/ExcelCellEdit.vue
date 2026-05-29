<template>
  <el-dialog
    :model-value="visible"
    title="编辑单元格"
    width="500px"
    @update:model-value="emit('close')"
    @close="emit('close')"
  >
    <div class="edit-info" v-if="rowData !== null">
      <p><strong>行 {{ rowIndex + 1 }}</strong> · 字段: {{ field }}</p>
      <p class="edit-errors" v-if="currentErrors.length">
        <span v-for="err in currentErrors" :key="err.message" class="edit-error-item">
          ⚠ {{ err.message }}
        </span>
      </p>
    </div>

    <el-input
      v-model="editValue"
      type="textarea"
      :rows="3"
      placeholder="输入新值"
      @keydown.esc="emit('close')"
    />

    <template #footer>
      <el-button @click="emit('close')">取消</el-button>
      <el-button type="primary" @click="save">保存</el-button>
    </template>
  </el-dialog>
</template>

<script setup lang="ts">
import { ref, watch } from 'vue'
import type { CellError } from '../types'

const props = defineProps<{
  visible: boolean
  rowIndex: number
  field: string
  value: any
  currentErrors: CellError[]
  rowData: Record<string, any> | null
}>()

const emit = defineEmits<{
  (e: 'save', rowIndex: number, field: string, value: any): void
  (e: 'close'): void
}>()

const editValue = ref('')

watch(() => props.visible, (v) => {
  if (v) editValue.value = String(props.value ?? '')
})

function save() {
  emit('save', props.rowIndex, props.field, editValue.value)
}
</script>

<style scoped>
.edit-info { margin-bottom: 12px; }
.edit-errors { margin: 4px 0; }
.edit-error-item {
  display: block;
  color: #dc2626;
  font-size: 13px;
}
</style>
