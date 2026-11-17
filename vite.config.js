import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';
import { VitePWA } from 'vite-plugin-pwa';

export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    VitePWA({
      registerType: 'autoUpdate',
      manifestFilename: 'manifest.json',
      includeAssets: ['icon.svg'],
      manifest: {
        name: 'מעקב מתקני מים - AquaOps PRO',
        short_name: 'מעקב מים',
        description: 'מערכת ניהול ומעקב מתקני מים ומיכלים',
        start_url: '/',
        display: 'standalone',
        background_color: '#f8fafc',
        theme_color: '#4f46e5',
        dir: 'rtl',
        lang: 'he',
        icons: [
          {
            src: '/icon.svg',
            sizes: 'any',
            type: 'image/svg+xml',
            purpose: 'any maskable'
          }
        ]
      },
      workbox: {
        globPatterns: ['**/*.{js,css,html,svg,png,ico,woff2}'],
        runtimeCaching: [
          {
            urlPattern: /^https:\/\/cdn\.sheetjs\.com\/.*/i,
            handler: 'CacheFirst',
            options: {
              cacheName: 'sheetjs-xlsx-cache',
              expiration: {
                maxEntries: 5,
                maxAgeSeconds: 60 * 60 * 24 * 365
              },
              cacheableResponse: {
                statuses: [0, 200]
              }
            }
          }
        ]
      }
    })
  ]
});