import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath, URL } from 'node:url'

/** Same-origin proxies (LAN preview/dev without CORS pain). */
const sharedProxy = {
  '/google-sheets': {
    target: 'https://docs.google.com',
    changeOrigin: true,
    secure: true,
    rewrite: (path: string) => path.replace(/^\/google-sheets/, ''),
  },
  '/api': {
    target: 'http://127.0.0.1:4000',
    changeOrigin: true,
    rewrite: (path: string) => path.replace(/^\/api/, ''),
  },
} as const

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    host: true,
    port: 5173,
    proxy: { ...sharedProxy },
  },
  preview: {
    host: true,
    port: 4173,
    proxy: { ...sharedProxy },
  },
})
