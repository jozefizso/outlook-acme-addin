import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import officeAddin from 'vite-plugin-office-addin'
import * as certs from 'office-addin-dev-certs'

export default defineConfig(async ({ command, mode, isSsrBuild, isPreview }) => {
  const env = loadEnv(mode, process.cwd(), '')

  return {
    root: 'src',
    publicDir: '../public',
    envDir: '../',
    build: {
      manifest: true,
      outDir: '../dist',
      rollupOptions: {
        input: {
          'index': 'src/index.html',
          'taskpane': 'src/taskpane.html',
        }
      }
    },
    plugins: [
      react(),
      officeAddin()
    ],
    server: {
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Document-Policy': 'js-profiling',
      },
      https: mode === 'development' ? await certs.getHttpsServerOptions() : undefined,
      port: env.PORT || 3000
    }
  }
})
