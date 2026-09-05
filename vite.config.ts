/// <reference types="vitest/config" />
import { defineConfig, type UserConfig } from 'vite'
import type { InlineConfig } from 'vitest'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { execFileSync } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { networkInterfaces } from 'node:os'
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

const lanIpv4 = () => {
  try {
    return Object.values(networkInterfaces())
      .flat()
      .filter((iface): iface is NonNullable<typeof iface> =>
        Boolean(
          iface &&
            !iface.internal &&
            (iface.family === 'IPv4' || String(iface.family) === '4'),
        ),
      )
      .map((iface) => iface.address)
  } catch {
    return []
  }
}

const mkcertBin = () => {
  const candidates = [
    'mkcert',
    `${process.env.HOME || ''}/.local/bin/mkcert`,
    '/opt/homebrew/bin/mkcert',
    '/usr/local/bin/mkcert',
  ]
  for (const candidate of candidates) {
    try {
      if (candidate === 'mkcert') {
        return execFileSync('which', ['mkcert'], { encoding: 'utf8' }).trim()
      }
      if (existsSync(candidate)) return candidate
    } catch {
      /* try next */
    }
  }
  return ''
}

/**
 * Trusted cert if `mkcert` is installed (`npm run cert:trust`).
 * Otherwise a self-signed fallback (browser: Not Secure).
 * Set VITE_HTTP=1 to keep plain HTTP.
 */
const loadLanHttps = () => {
  const dir = fileURLToPath(new URL('./.cert', import.meta.url))
  const keyFile = `${dir}/key.pem`
  const certFile = `${dir}/cert.pem`
  const stampFile = `${dir}/san.txt`
  const hosts = [...new Set(['localhost', '127.0.0.1', '::1', ...lanIpv4()])]
  const stamp = hosts.slice().sort().join(',')
  const stampOk =
    existsSync(stampFile) && readFileSync(stampFile, 'utf8').trim() === stamp
  if (!existsSync(keyFile) || !existsSync(certFile) || !stampOk) {
    mkdirSync(dir, { recursive: true })
    const mkcert = mkcertBin()
    if (mkcert) {
      execFileSync(mkcert, ['-key-file', keyFile, '-cert-file', certFile, ...hosts], {
        stdio: 'ignore',
      })
    } else {
      const san = hosts
        .map((host) =>
          /^\d{1,3}(?:\.\d{1,3}){3}$/.test(host)
            ? `IP:${host}`
            : host.includes(':')
              ? `IP:${host}`
              : `DNS:${host}`,
        )
        .join(',')
      execFileSync(
        'openssl',
        [
          'req',
          '-x509',
          '-newkey',
          'rsa:2048',
          '-sha256',
          '-days',
          '825',
          '-nodes',
          '-keyout',
          keyFile,
          '-out',
          certFile,
          '-subj',
          '/CN=army-grid-lan',
          '-addext',
          `subjectAltName=${san}`,
        ],
        { stdio: 'ignore' },
      )
    }
    writeFileSync(stampFile, stamp)
  }
  return {
    key: readFileSync(keyFile),
    cert: readFileSync(certFile),
  }
}

const https = process.env.VITE_HTTP === '1' ? undefined : loadLanHttps()
const test = {
  environment: 'node',
  include: ['src/**/*.test.ts'],
} satisfies InlineConfig

// https://vite.dev/config/
const config = {
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
  server: {
    // Explicit IPv4 bind so phones/other PCs on Wi‑Fi can reach the Mac.
    host: '0.0.0.0',
    port: 5173,
    strictPort: true,
    allowedHosts: true,
    https,
    proxy: { ...sharedProxy },
  },
  preview: {
    host: '0.0.0.0',
    port: 4173,
    strictPort: true,
    allowedHosts: true,
    https,
    proxy: { ...sharedProxy },
  },
  test,
} satisfies UserConfig & { test: InlineConfig }

export default defineConfig(config)
