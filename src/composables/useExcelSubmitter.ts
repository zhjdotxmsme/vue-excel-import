import type { SubmitResult, SubmitProgress } from '../types'

interface ExtendedSubmitConfig {
  submitApi: string | ((data: Record<string, any>[]) => Promise<any>)
  rollbackApi?: string | ((batchId: string) => Promise<any>)
  batchSize?: number
  onProgress?: (progress: SubmitProgress) => void
}

export async function useExcelSubmitter(
  rows: Record<string, any>[],
  config: ExtendedSubmitConfig
): Promise<SubmitResult> {
  const { submitApi, rollbackApi } = config
  const batchSize = config.batchSize ?? 500
  const totalBatches = Math.ceil(rows.length / batchSize)
  const completedBatches: any[] = []

  async function sendBatch(batch: Record<string, any>[]): Promise<any> {
    if (typeof submitApi === 'function') {
      return submitApi(batch)
    }
    const response = await fetch(submitApi, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(batch)
    })
    if (!response.ok) throw new Error(`HTTP ${response.status}`)
    return response.json()
  }

  async function rollback(): Promise<boolean> {
    if (!rollbackApi) return false
    try {
      if (typeof rollbackApi === 'function') {
        await rollbackApi(String(Date.now()))
      } else {
        await fetch(rollbackApi, { method: 'POST' })
      }
      return true
    } catch {
      return false
    }
  }

  for (let i = 0; i < totalBatches; i++) {
    const start = i * batchSize
    const batch = rows.slice(start, start + batchSize)
    const batchNum = i + 1

    try {
      const result = await sendBatch(batch)
      completedBatches.push(result)

      config.onProgress?.({
        current: Math.min((i + 1) * batchSize, rows.length),
        total: rows.length,
        batch: batchNum,
        totalBatches
      })
    } catch (err: any) {
      const rollbackSuccess = await rollback()
      return {
        success: false,
        totalBatches,
        completedBatches: i,
        failedBatch: batchNum,
        rollbackSuccess,
        error: err.message || '提交失败'
      }
    }
  }

  return {
    success: true,
    totalBatches,
    completedBatches: totalBatches,
    rollbackSuccess: true
  }
}
