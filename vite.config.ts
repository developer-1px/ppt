import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'

const canvasCommandPaletteItemsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/command-palette/CanvasCommandPaletteItems.ts', import.meta.url))
const canvasKeyboardCommandDispatchEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardCommandDispatch.ts', import.meta.url))
const canvasKeyboardCommandShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardCommandShortcuts.ts', import.meta.url))
const canvasKeyboardNudgeShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardNudgeShortcuts.ts', import.meta.url))
const canvasKeyboardViewportDispatchEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardViewportDispatch.ts', import.meta.url))
const canvasKeyboardViewportShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardViewportShortcuts.ts', import.meta.url))
const canvasMenuRovingFocusEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/toolbar/CanvasMenuRovingFocus.ts', import.meta.url))
const canvasMinimapModelEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/minimap/CanvasMinimapModel.ts', import.meta.url))
const canvasRadioGroupEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/radio/CanvasRadioGroup.ts', import.meta.url))
const canvasToolbarRovingFocusEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/toolbar/CanvasToolbarRovingFocus.ts', import.meta.url))
const canvasCoreEntry = fileURLToPath(new URL('../canvas/src/canvas/core/index.ts', import.meta.url))
const canvasFoundationEntry = fileURLToPath(new URL('../canvas/src/canvas/foundation/index.ts', import.meta.url))
const canvasEngineEntry = fileURLToPath(new URL('../canvas/src/canvas/engine/index.ts', import.meta.url))
const slideEditAffordanceEntry = fileURLToPath(new URL('../canvas/packages/slide-edit-affordance/src/index.ts', import.meta.url))

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: 'canvas/app/command-palette-items', replacement: canvasCommandPaletteItemsEntry },
      { find: 'canvas/app/keyboard-command-dispatch', replacement: canvasKeyboardCommandDispatchEntry },
      { find: 'canvas/app/keyboard-command-shortcuts', replacement: canvasKeyboardCommandShortcutsEntry },
      { find: 'canvas/app/keyboard-nudge-shortcuts', replacement: canvasKeyboardNudgeShortcutsEntry },
      { find: 'canvas/app/keyboard-viewport-dispatch', replacement: canvasKeyboardViewportDispatchEntry },
      { find: 'canvas/app/keyboard-viewport-shortcuts', replacement: canvasKeyboardViewportShortcutsEntry },
      { find: 'canvas/app/menu-roving-focus', replacement: canvasMenuRovingFocusEntry },
      { find: 'canvas/app/minimap-model', replacement: canvasMinimapModelEntry },
      { find: 'canvas/app/radio-group', replacement: canvasRadioGroupEntry },
      { find: 'canvas/app/toolbar-roving-focus', replacement: canvasToolbarRovingFocusEntry },
      { find: 'canvas/core', replacement: canvasCoreEntry },
      { find: 'canvas/foundation', replacement: canvasFoundationEntry },
      { find: 'canvas/engine', replacement: canvasEngineEntry },
      { find: '@interactive-os/slide-edit-affordance', replacement: slideEditAffordanceEntry },
    ],
    dedupe: ['react', 'react-dom', 'zod'],
  },
})
