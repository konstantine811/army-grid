import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath, URL } from 'node:url'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    proxy: {
      // Avoid browser CORS when loading public Google Sheets CSV exports.
      '/google-sheets': {
        target: 'https://docs.google.com',
        changeOrigin: true,
        secure: true,
        rewrite: (path) => path.replace(/^\/google-sheets/, ''),
      },
    },
  },
})
