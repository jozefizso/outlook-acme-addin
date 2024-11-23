import { defineConfig, loadEnv } from 'vite'
import officeAddin from 'vite-plugin-office-addin'
import * as certs from 'office-addin-dev-certs'

export default defineConfig(async ({ command, mode, isSsrBuild, isPreview }) => {
  const env = loadEnv(mode, process.cwd(), '')

  return {
    root: 'src',
    publicDir: '../public',
    build: {
      manifest: true,
      outDir: '../dist',
      rollupOptions: {
        input: {
          'taskpane': 'src/taskpane.html',
        }
      }
    },
    plugins: [
      officeAddin()
    ],
    server: {
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      https: mode === 'development' ? await certs.getHttpsServerOptions() : undefined,
      port: env.PORT || 3000
    }
  }
})
