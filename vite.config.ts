import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    host: '0.0.0.0',
    port: 5181,
    strictPort: true,
    proxy: {
      '/api': {
        // Явный IPv4: в средах с предпочтением IPv6 localhost может дать ECONNREFUSED.
        target: 'http://127.0.0.1:8787',
        changeOrigin: true,
      },
    },
  },
})
