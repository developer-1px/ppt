import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'

const zodCrudEntry = fileURLToPath(import.meta.resolve('zod-crud'))
const zodCrudReactEntry = fileURLToPath(import.meta.resolve('zod-crud/react'))

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: 'zod-crud/react', replacement: zodCrudReactEntry },
      { find: 'zod-crud', replacement: zodCrudEntry },
    ],
    dedupe: ['react', 'react-dom', 'zod'],
  },
})
