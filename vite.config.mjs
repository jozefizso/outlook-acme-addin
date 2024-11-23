import { defineConfig, loadEnv } from 'vite'

export default defineConfig(({ command, mode, isSsrBuild, isPreview }) => {
  const env = loadEnv(mode, process.cwd(), '')

  return {
    root: 'src',
    build: {
      manifest: true,
      outDir: '../dist',
      rollupOptions: {
        input: {
          'taskpane': 'src/taskpane.html',
        }
      }
    },
    server: {
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      port: env.PORT || 3000
    }
  }
})
