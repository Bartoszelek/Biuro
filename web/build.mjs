import * as esbuild from 'esbuild'

// Paczki dla WebView – wszystko lokalnie, bez internetu.
await esbuild.build({
  entryPoints: { sheet: 'src/sheet.js', doc: 'src/doc.js' },
  bundle: true,
  minify: true,
  format: 'esm',
  platform: 'browser',
  target: ['chrome100'],
  outdir: '../app/src/main/assets/www',
  loader: { '.woff': 'file', '.woff2': 'file', '.ttf': 'file', '.eot': 'file', '.svg': 'file', '.png': 'file', '.gif': 'file', '.jpg': 'file' },
  assetNames: 'a/[name]-[hash]',
  define: { 'process.env.NODE_ENV': '"production"', global: 'globalThis' },
  logLevel: 'info',
})
