import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import dts from 'vite-plugin-dts'
import { resolve } from 'path'

export default defineConfig({
  plugins: [vue(), dts({ insertTypesEntry: true })],
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
        exports: 'named',
        globals: {
          vue: 'Vue',
          'element-plus': 'ElementPlus',
          exceljs: 'ExcelJS'
        }
      }
    }
  }
})