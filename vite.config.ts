import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { resolve } from 'path'

export default defineConfig({
  plugins: [vue()],
  resolve: { alias: { '@': resolve(__dirname, 'src') } },
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'VueExcelImport',
      fileName: 'vue-excel-import'
    },
    rollupOptions: {
      external: ['vue', 'element-plus', 'exceljs'],
      output: {
        globals: {
          vue: 'Vue',
          'element-plus': 'ElementPlus',
          exceljs: 'ExcelJS'
        }
      }
    }
  }
})