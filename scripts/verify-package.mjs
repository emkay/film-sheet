// Install the packed package into a throwaway project and use it the way a
// consumer would: import the engine in Node, and typecheck both entry points
// with TypeScript 5 under `bundler` and strict `node16` module resolution.
// Run after `npm run build:lib`.
import { execFileSync } from 'node:child_process'
import { mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const run = (cmd, args, cwd) => execFileSync(cmd, args, { cwd, stdio: ['ignore', 'pipe', 'inherit'], encoding: 'utf8' })
const dir = mkdtempSync(join(tmpdir(), 'film-sheet-consumer-'))

try {
  run('npm', ['pack', '--pack-destination', dir, '--silent'], process.cwd())
  const tarball = readdirSync(dir).find((f) => f.endsWith('.tgz'))
  if (!tarball) throw new Error('npm pack produced no tarball')

  writeFileSync(join(dir, 'package.json'), JSON.stringify({ name: 'consumer', private: true, type: 'module' }))
  run('npm', ['install', '--silent', '--no-audit', '--no-fund', `./${tarball}`, 'lit', '@mk/film', 'typescript@5'], dir)

  writeFileSync(join(dir, 'engine.mjs'), `
import { Workbook, registerFunction, toNumber, isError } from '@mk/film-sheet/engine'
registerFunction('DOUBLE', { min: 1, max: 1, fn: ([x]) => { const n = toNumber(x); return isError(n) ? n : n * 2 } })
const wb = new Workbook()
const id = wb.sheets[0].id
wb.setCell(id, 0, 0, '21')
wb.setCell(id, 1, 0, '=DOUBLE(A1)')
if (wb.getValue(id, 1, 0) !== 42) throw new Error('engine gave ' + wb.getValue(id, 1, 0))
`)
  run('node', ['engine.mjs'], dir)
  console.log('✓ engine runs in Node')

  writeFileSync(join(dir, 'app.ts'), `
import { FilmSheet, type WorkbookJSON, type SheetSelectDetail } from '@mk/film-sheet'
import { Workbook, registerFunction, type Scalar } from '@mk/film-sheet/engine'

const sheet: FilmSheet = document.createElement('film-sheet')
sheet.readonly = true
sheet.chrome = 'minimal'
const data: WorkbookJSON = { version: 1, sheets: [{ name: 'S', cells: [[0, 0, '1']] }] }
sheet.load(data)
sheet.addEventListener('film-sheet-select', (e: CustomEvent<SheetSelectDetail>) => e.detail.range)
const v: Scalar = new Workbook().getValue(1, 0, 0)
registerFunction('X', { fn: () => v })
// @ts-expect-error engine internals are not part of the public types
sheet.workbook?.missing
`)
  for (const [module, resolution] of [['ESNext', 'bundler'], ['Node16', 'node16']]) {
    writeFileSync(join(dir, 'tsconfig.json'), JSON.stringify({
      compilerOptions: { target: 'ES2022', module, moduleResolution: resolution, strict: true, noEmit: true, lib: ['ES2022', 'DOM'], skipLibCheck: false, types: [] },
      files: ['app.ts']
    }))
    run(join(dir, 'node_modules/.bin/tsc'), ['-p', 'tsconfig.json'], dir)
    console.log(`✓ types check with TypeScript 5 (${resolution})`)
  }
} finally {
  rmSync(dir, { recursive: true, force: true })
}
