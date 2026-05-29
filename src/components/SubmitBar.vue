<template>
  <div class="submit-bar" v-if="total > 0">
    <div class="stats">
      <span class="stat-item">总行数: <strong>{{ total }}</strong></span>
      <span class="stat-divider">|</span>
      <span class="stat-item stat-pass">通过: <strong>{{ validCount }}</strong></span>
      <span class="stat-divider">|</span>
      <span class="stat-item" :class="{ 'stat-fail': errorCount > 0 }">
        错误: <strong>{{ errorCount }}</strong>
      </span>
    </div>

    <div class="actions">
      <el-progress
        v-if="submitting"
        :percentage="progressPercent"
        :stroke-width="16"
        :text-inside="true"
        style="width: 200px; margin-right: 12px;"
      />

      <el-button
        type="primary"
        size="large"
        :disabled="!canSubmit"
        :loading="submitting"
        @click="emit('submit')"
      >
        {{ submitting ? `提交中 (${progress.current}/${progress.total})` : submitText }}
      </el-button>

      <el-button
        v-if="errorCount > 0 && !submitting"
        size="large"
        @click="emit('fix-errors')"
      >
        修正后提交
      </el-button>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue'
import type { SubmitProgress } from '../types'

const props = defineProps<{
  total: number
  validCount: number
  errorCount: number
  submitting: boolean
  progress: SubmitProgress
}>()

const emit = defineEmits<{
  (e: 'submit'): void
  (e: 'fix-errors'): void
}>()

const canSubmit = computed(() => props.errorCount === 0 && props.total > 0 && !props.submitting)

const submitText = computed(() => {
  if (props.total === 0) return '导入数据'
  const batches = Math.ceil(props.total / 500)
  return `导入数据（${props.total}条，分${batches}批）`
})

const progressPercent = computed(() => {
  if (props.progress.total === 0) return 0
  return Math.round((props.progress.current / props.progress.total) * 100)
})
</script>

<style scoped>
.submit-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 12px 16px;
  background: #fafafa;
  border: 1px solid #e4e7ed;
  border-radius: 8px;
  flex-wrap: wrap;
  gap: 12px;
}

.stats { display: flex; align-items: center; gap: 8px; font-size: 14px; }
.stat-divider { color: #dcdfe6; }
.stat-pass strong { color: #67c23a; }
.stat-fail strong { color: #f56c6c; }

.actions { display: flex; align-items: center; }
</style>
