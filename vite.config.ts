import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'

const canvasCoreEntry = fileURLToPath(new URL('../canvas/src/canvas/core/index.ts', import.meta.url))
const canvasFoundationEntry = fileURLToPath(new URL('../canvas/src/canvas/foundation/index.ts', import.meta.url))
const canvasEngineEntry = fileURLToPath(new URL('../canvas/src/canvas/engine/index.ts', import.meta.url))
const slideEditAffordanceEntry = fileURLToPath(new URL('../canvas/packages/slide-edit-affordance/src/index.ts', import.meta.url))

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: 'canvas/core', replacement: canvasCoreEntry },
      { find: 'canvas/foundation', replacement: canvasFoundationEntry },
      { find: 'canvas/engine', replacement: canvasEngineEntry },
      { find: '@interactive-os/slide-edit-affordance', replacement: slideEditAffordanceEntry },
    ],
    dedupe: ['react', 'react-dom', 'zod'],
  },
})
