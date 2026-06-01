# AGENTS.md

## Build & Test

- `npm run build` = `vue-tsc --noEmit && vite build` (type-check first, then build)
- `npm test` = `vitest run` (jsdom env, globals: true, tests in `src/**/*.test.ts`)
- `npm run dev` = `vite demo` (serves `demo/` as an SPA, not library dev mode)
- `npm run test:watch` for watch mode
- Single test: `npx vitest run src/composables/useExcelParser.test.ts`

## Dependencies

- **Runtime**: only `hucre` (zero transitive deps, ~18KB)
- **Peer**: `vue@^3.4`, `element-plus@^2.8` (consumers install these)
- **Dev**: `exceljs` only for test fixtures (`makeExcelFile` helper) — never imported in production code
- Import hucre subpath: `import { readXlsx } from 'hucre/xlsx'` or `import { writeXlsx } from 'hucre/xlsx'`

## Architecture

- **Entry**: `src/index.ts` → default export = `ExcelImport` (Vue SFC), named export = `{ ExcelImport }`, re-exports all types
- **Parsing**: Web Worker (`excelParser.worker.ts`) using hucre, sends batched chunks via `postMessage`; falls back to main-thread when Worker unavailable
- **Template generation**: `useExcelTemplate` → hucre `writeXlsx` → Blob
- **Validation**: sync + async validator support, `useExcelValidator` auto-detects async custom validators
- **Components**: `ExcelImport` orchestrates everything; sub-components: `ExcelUploader`, `ExcelPreview`, `ExcelCellEdit`, `SubmitBar`

## Multi-level Header Support

- `ColumnConfig.children` enables nested group headers
- `detectHeaderTree()` reads merged cells + row structure from raw Excel data → `ColumnConfig[]` tree
- `matchColumnTree()` aligns user config tree against detected Excel tree using label matching
- `stripAsteriskNode()` strips trailing `*` from template-generated headers so matching is template-compatible
- Utilities: `flattenColumns()` (tree → leaf list), `setNested()` / `getNested()` (dot-path access for nested fields)

## Known Stale Config

- `vite.config.ts` lists `exceljs` in `rollupOptions.external` — no longer imported at runtime, harmless but stale
- `.github/workflows/webpack.yml` runs `npx webpack` — should be `npm run build` instead

## Test Conventions

- Tests co-located with source files (`*.test.ts` next to `*.ts`)
- Template tests use hucre `readXlsx` for readback verification (data validations written but may not roundtrip detectably)
- Parser tests use `makeExcelFile` (ExcelJS) to create test fixtures, then parse them via hucre
- File reading in tests uses a `blobToArrayBuffer` helper since jsdom lacks native `Blob.arrayBuffer()`
