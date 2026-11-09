import { build } from 'esbuild'

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
