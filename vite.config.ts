import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// https://vite.dev/config/
export default defineConfig({
  // GitHub Pages sirve el sitio bajo /Whiteboard/ (nombre del repo)
  base: '/Whiteboard/',
  plugins: [react()],
})
