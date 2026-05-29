import { describe, it, expect, vi, beforeEach } from 'vitest'
import { useExcelSubmitter } from './useExcelSubmitter'

describe('useExcelSubmitter', () => {
  beforeEach(() => {
    vi.restoreAllMocks()
  })

  it('submits all rows in one batch when under batch size', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const result = await useExcelSubmitter(
      [{ name: 'A' }, { name: 'B' }],
      { submitApi: fn, batchSize: 500 }
    )
    expect(fn).toHaveBeenCalledTimes(1)
    expect(fn).toHaveBeenCalledWith([{ name: 'A' }, { name: 'B' }])
    expect(result.success).toBe(true)
    expect(result.totalBatches).toBe(1)
  })

  it('splits into multiple batches when over batch size', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const data = Array.from({ length: 1200 }, (_, i) => ({ id: i + 1 }))
    const result = await useExcelSubmitter(data, { submitApi: fn, batchSize: 500 })
    expect(fn).toHaveBeenCalledTimes(3)
    expect(fn.mock.calls[0][0]).toHaveLength(500)
    expect(fn.mock.calls[1][0]).toHaveLength(500)
    expect(fn.mock.calls[2][0]).toHaveLength(200)
    expect(result.success).toBe(true)
    expect(result.totalBatches).toBe(3)
  })

  it('calls rollback on failure and returns rollbackSuccess', async () => {
    const submitFn = vi.fn()
      .mockResolvedValueOnce({ ok: true })
      .mockRejectedValueOnce(new Error('batch 2 failed'))
    const rollbackFn = vi.fn().mockResolvedValue({ ok: true })
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i + 1 }))

    const result = await useExcelSubmitter(data, {
      submitApi: submitFn,
      rollbackApi: rollbackFn,
      batchSize: 500
    })

    expect(submitFn).toHaveBeenCalledTimes(2)
    expect(rollbackFn).toHaveBeenCalledTimes(1)
    expect(result.success).toBe(false)
    expect(result.failedBatch).toBe(2)
    expect(result.rollbackSuccess).toBe(true)
  })

  it('reports progress via onProgress callback', async () => {
    const fn = vi.fn().mockResolvedValue({ ok: true })
    const progresses: any[] = []
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i + 1 }))

    await useExcelSubmitter(data, {
      submitApi: fn,
      batchSize: 500,
      onProgress: (p) => progresses.push(p)
    })

    expect(progresses).toHaveLength(2)
    expect(progresses[0]).toMatchObject({ batch: 1, totalBatches: 2 })
    expect(progresses[1]).toMatchObject({ batch: 2, totalBatches: 2 })
  })

  it('calls submitApi as string URL via fetch', async () => {
    globalThis.fetch = vi.fn().mockResolvedValue({
      ok: true,
      json: () => Promise.resolve({})
    })
    const data = [{ name: 'A' }]
    const result = await useExcelSubmitter(data, { submitApi: '/api/import', batchSize: 500 })
    expect(globalThis.fetch).toHaveBeenCalledWith('/api/import', expect.objectContaining({
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify([{ name: 'A' }])
    }))
    expect(result.success).toBe(true)
  })

  it('marks rollbackSuccess as false when rollback itself fails', async () => {
    const submitFn = vi.fn()
      .mockResolvedValueOnce({ ok: true })
      .mockRejectedValueOnce(new Error('failed'))
    const rollbackFn = vi.fn().mockRejectedValue(new Error('rollback also failed'))
    const data = Array.from({ length: 1000 }, (_, i) => ({ id: i + 1 }))

    const result = await useExcelSubmitter(data, {
      submitApi: submitFn,
      rollbackApi: rollbackFn,
      batchSize: 500
    })

    expect(result.success).toBe(false)
    expect(result.rollbackSuccess).toBe(false)
  })
})
