import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Hemos removido el plugin experimental de Tailwind v4 que rompía la plataforma
export default defineConfig({
  plugins: [react()],
})