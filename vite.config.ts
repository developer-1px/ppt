import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'

const canvasCommandPaletteItemsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/command-palette/CanvasCommandPaletteItems.ts', import.meta.url))
const canvasImageClipboardEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/image-io/CanvasImageClipboard.ts', import.meta.url))
const canvasImageImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/image-io/CanvasImageImport.ts', import.meta.url))
const canvasInlineEditDomEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/editing/text-editor/CanvasInlineEditDom.ts', import.meta.url))
const canvasKeyboardSystemShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardSystemShortcuts.ts', import.meta.url))
const canvasMediaImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/media-import/CanvasMediaImport.ts', import.meta.url))
const canvasMinimapModelEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/minimap/CanvasMinimapModel.ts', import.meta.url))
const canvasPastePositionEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/commands/CanvasPastePosition.ts', import.meta.url))
const canvasPointerClickMemoryEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerClickMemory.ts', import.meta.url))
const canvasPointerStartSessionEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerStartSession.ts', import.meta.url))
const canvasAppEntry = fileURLToPath(new URL('../canvas/src/canvas/app/index.ts', import.meta.url))
const pptCanvasAppCompatEntry = fileURLToPath(new URL('./src/pptCanvasAppCompat.ts', import.meta.url))
const pptCanvasHostCompatEntry = fileURLToPath(new URL('./src/pptCanvasHostCompat.ts', import.meta.url))
const pptCanvasImageImportCompatEntry = fileURLToPath(new URL('./src/pptCanvasImageImportCompat.ts', import.meta.url))
const pptCanvasMediaImportCompatEntry = fileURLToPath(new URL('./src/pptCanvasMediaImportCompat.ts', import.meta.url))
const canvasCoreEntry = fileURLToPath(new URL('../canvas/src/canvas/core/index.ts', import.meta.url))
const canvasFoundationEntry = fileURLToPath(new URL('../canvas/src/canvas/foundation/index.ts', import.meta.url))
const canvasEngineEntry = fileURLToPath(new URL('../canvas/src/canvas/engine/index.ts', import.meta.url))
const canvasHostEntry = fileURLToPath(new URL('../canvas/src/canvas/host/index.ts', import.meta.url))
const canvasRendererEntry = fileURLToPath(new URL('../canvas/src/canvas/renderer/index.ts', import.meta.url))
const canvasSvgDrawingPrimitivesEntry = fileURLToPath(new URL('../canvas/src/canvas/renderer/svg/CanvasSvgDrawingPrimitives.ts', import.meta.url))
const domEditAffordanceMetadataEntry = fileURLToPath(new URL('../canvas/packages/dom-edit-affordance/src/metadata.ts', import.meta.url))
const slideEditAffordanceEntry = fileURLToPath(new URL('../canvas/packages/slide-edit-affordance/src/index.ts', import.meta.url))

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: 'canvas/app/command-palette-items', replacement: canvasCommandPaletteItemsEntry },
      { find: 'canvas/app/image-clipboard', replacement: canvasImageClipboardEntry },
      { find: 'canvas/app/image-import-public', replacement: canvasImageImportEntry },
      { find: 'canvas/app/image-import', replacement: pptCanvasImageImportCompatEntry },
      { find: 'canvas/app/inline-edit-dom', replacement: canvasInlineEditDomEntry },
      { find: 'canvas/app/keyboard-system-shortcuts', replacement: canvasKeyboardSystemShortcutsEntry },
      { find: 'canvas/app/media-import-public', replacement: canvasMediaImportEntry },
      { find: 'canvas/app/media-import', replacement: pptCanvasMediaImportCompatEntry },
      { find: 'canvas/app/minimap-model', replacement: canvasMinimapModelEntry },
      { find: 'canvas/app/paste-position', replacement: canvasPastePositionEntry },
      { find: 'canvas/app/pointer-click-memory', replacement: canvasPointerClickMemoryEntry },
      { find: 'canvas/app/pointer-start-session', replacement: canvasPointerStartSessionEntry },
      { find: 'canvas/app-public', replacement: canvasAppEntry },
      { find: 'canvas/app', replacement: pptCanvasAppCompatEntry },
      { find: 'canvas/core', replacement: canvasCoreEntry },
      { find: 'canvas/foundation', replacement: canvasFoundationEntry },
      { find: 'canvas/engine', replacement: canvasEngineEntry },
      { find: 'canvas/host-public', replacement: canvasHostEntry },
      { find: 'canvas/host', replacement: pptCanvasHostCompatEntry },
      { find: 'canvas/renderer/svg-drawing-primitives', replacement: canvasSvgDrawingPrimitivesEntry },
      { find: 'canvas/renderer', replacement: canvasRendererEntry },
      { find: '@interactive-os/dom-edit-affordance/metadata', replacement: domEditAffordanceMetadataEntry },
      { find: '@interactive-os/slide-edit-affordance', replacement: slideEditAffordanceEntry },
    ],
    dedupe: ['react', 'react-dom', 'zod'],
  },
})
