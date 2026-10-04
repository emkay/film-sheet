import { resolve } from 'node:path'
import { defineConfig } from 'vitest/config'

// `vite build` builds the demo site; `vite build --mode lib` builds the package.
export default defineConfig(({ mode }) => ({
  build: mode === 'lib'
    ? {
        target: 'es2022',
        outDir: 'dist',
        emptyOutDir: true,
        sourcemap: true,
        lib: {
          entry: { index: resolve(import.meta.dirname, 'src/index.ts'), engine: resolve(import.meta.dirname, 'src/engine/index.ts') },
          formats: ['es']
        },
        rollupOptions: {
          // Peer dependencies: the host app provides one copy of Lit and Film.
          external: [/^lit($|\/)/, /^@lit\//, /^@mk\/film($|\/)/],
          output: { preserveModules: true, preserveModulesRoot: 'src', entryFileNames: '[name].js' }
        }
      }
    : { target: 'es2022', outDir: 'dist-demo' },
  test: {
    include: ['test/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      // Canvas/DOM wiring can't run under Node; e2e/ covers it in a real browser.
      exclude: ['src/index.ts', 'src/engine/index.ts', 'src/ui/controller.ts', 'src/ui/grid.ts', 'src/ui/editor.ts', 'src/ui/film-sheet.ts', 'src/ui/template.ts'],
      reporter: ['text', 'html'],
      thresholds: { lines: 90, statements: 85, functions: 90, branches: 70 }
    }
  }
}))
