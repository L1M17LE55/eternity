import { build } from 'esbuild'
import { readFile } from 'node:fs/promises'

const metadata = await readFile('src/metadata.json', 'utf8')

await build({
  entryPoints: ['src/app.js'],
  bundle: true,
  outfile: 'dist/app.bundle.js',
  format: 'iife',
  globalName: 'App',
  platform: 'neutral',
  target: 'es2020',
  loader: {
    '.html': 'text'
  },
  footer: {
    js: 'return App;'
  }
})
