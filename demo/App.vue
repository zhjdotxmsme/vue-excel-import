<template>
  <div class="demo-container">
    <h1 class="demo-title">Excel 导入组件 Demo</h1>
    <p class="demo-desc">配置列映射和校验规则，上传 Excel 文件验证解析效果</p>

    <ExcelImport
      :columns="columns"
      :submit-api="mockSubmitApi"
      :rollback-api="mockRollbackApi"
      :batch-size="500"
      :show-template-button="true"
      template-description="请按照模板格式填写数据，标 * 为必填项"
      template-example-rows="1"
      @parsed="onParsed"
      @progress="onProgress"
      @success="onSuccess"
      @error="onError"
    />
  </div>
</template>

<script setup lang="ts">
import { ElNotification } from 'element-plus'
import type { ColumnConfig, ParseResult, SubmitProgress, SubmitResult } from '../src/types'
import ExcelImport from '../src/components/ExcelImport.vue'

// ========== 列配置（开发者初始化配置） ==========
const columns: ColumnConfig[] = [
  {
    label: '姓名',
    field: 'userName',
    type: 'string',
    required: true,
    description: '员工姓名'
  },
  {
    label: '年龄',
    field: 'age',
    type: 'number',
    validators: [
      { type: 'range', min: 0, max: 150, message: '年龄必须在 0-150 之间' }
    ]
  },
  {
    label: '邮箱',
    field: 'email',
    type: 'string',
    required: true,
    validators: [
      { type: 'pattern', pattern: '^[\\w.-]+@[\\w.-]+\\.\\w+$', message: '邮箱格式不正确' }
    ]
  },
  {
    label: '部门',
    field: 'dept',
    type: 'string',
    validators: [
      { type: 'enum', enum: ['技术部', '市场部', '财务部', '人事部'], message: '请选择正确的部门' }
    ]
  },
  {
    label: '入职日期',
    field: 'hireDate',
    type: 'date',
    description: '格式: 2024-01-15'
  },
  {
    label: '在职',
    field: 'active',
    type: 'boolean'
  }
]

// ========== Mock API（模拟后端接口） ==========
function mockSubmitApi(data: Record<string, any>[]): Promise<any> {
  console.log('📤 提交数据:', data)
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve({ success: true, count: data.length })
    }, 500)
  })
}

function mockRollbackApi(batchId: string): Promise<any> {
  console.log('↩️ 回滚批次:', batchId)
  return Promise.resolve({ success: true })
}

// ========== 事件处理 ==========
function onParsed(result: ParseResult) {
  console.log('📊 解析完成:', result.totalRows, '行')
}

function onProgress(progress: SubmitProgress) {
  console.log(`⏳ 提交进度: ${progress.batch}/${progress.totalBatches}`)
}

function onSuccess(result: SubmitResult) {
  ElNotification({
    title: '导入成功',
    message: `成功导入 ${result.totalBatches} 批数据`,
    type: 'success',
    duration: 3000
  })
}

function onError(error: Error) {
  ElNotification({
    title: '导入失败',
    message: error.message,
    type: 'error',
    duration: 5000
  })
}
</script>

<style>
body {
  margin: 0;
  padding: 20px;
  background: #f5f7fa;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
}

.demo-container {
  max-width: 960px;
  margin: 0 auto;
  background: #fff;
  padding: 32px;
  border-radius: 12px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
}

.demo-title {
  font-size: 24px;
  font-weight: 600;
  margin: 0 0 4px 0;
  color: #1a1a2e;
}

.demo-desc {
  font-size: 14px;
  color: #909399;
  margin: 0 0 24px 0;
}
</style>
