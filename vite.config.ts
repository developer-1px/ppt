import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { fileURLToPath } from 'node:url'

const canvasCommandPaletteItemsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/command-palette/CanvasCommandPaletteItems.ts', import.meta.url))
const canvasEraserHitTestingEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasEraserHitTesting.ts', import.meta.url))
const canvasPointerDrawingEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerDrawing.ts', import.meta.url))
const canvasImageImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/image-io/CanvasImageImport.ts', import.meta.url))
const canvasInlineEditDomEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/editing/text-editor/CanvasInlineEditDom.ts', import.meta.url))
const canvasKeyboardCommandDispatchEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardCommandDispatch.ts', import.meta.url))
const canvasKeyboardCommandShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardCommandShortcuts.ts', import.meta.url))
const canvasKeyboardNudgeShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardNudgeShortcuts.ts', import.meta.url))
const canvasKeyboardToolShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardToolShortcutIntent.ts', import.meta.url))
const canvasKeyboardViewportDispatchEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardViewportDispatch.ts', import.meta.url))
const canvasKeyboardViewportShortcutsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/keyboard/CanvasKeyboardViewportShortcuts.ts', import.meta.url))
const canvasMenuRovingFocusEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/toolbar/CanvasMenuRovingFocus.ts', import.meta.url))
const canvasMediaImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/media-import/CanvasMediaImport.ts', import.meta.url))
const canvasMinimapModelEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/minimap/CanvasMinimapModel.ts', import.meta.url))
const canvasModalFocusLifecycleEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/modal/CanvasModalFocusLifecycle.ts', import.meta.url))
const canvasPastePositionEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/commands/CanvasPastePosition.ts', import.meta.url))
const canvasPointerClickMemoryEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerClickMemory.ts', import.meta.url))
const canvasPointerGeometryEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerGeometry.ts', import.meta.url))
const canvasPointerLaserEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerLaser.ts', import.meta.url))
const canvasPointerPanInteractionEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/pointer/CanvasPointerPanInteraction.ts', import.meta.url))
const canvasRadioGroupEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/radio/CanvasRadioGroup.ts', import.meta.url))
const canvasRichClipboardEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/commands/CanvasRichClipboardIO.ts', import.meta.url))
const canvasTabsRovingFocusEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/controls/tabs/CanvasTabsRovingFocus.ts', import.meta.url))
const canvasStageElementEntry = fileURLToPath(new URL('../canvas/src/canvas/app/rendering/stage/CanvasAppStageElement.ts', import.meta.url))
const canvasTableImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/table-import/CanvasTableImport.ts', import.meta.url))
const canvasTextPasteImportEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/text-paste-import/CanvasTextPasteImport.ts', import.meta.url))
const canvasToolbarRovingFocusEntry = fileURLToPath(new URL('../canvas/src/canvas/app/feature-packs/toolbar/CanvasToolbarRovingFocus.ts', import.meta.url))
const canvasViewportControlsEntry = fileURLToPath(new URL('../canvas/src/canvas/app/affordances/interaction/viewport/CanvasViewportControlExecution.ts', import.meta.url))
const canvasCoreEntry = fileURLToPath(new URL('../canvas/src/canvas/core/index.ts', import.meta.url))
const canvasFoundationEntry = fileURLToPath(new URL('../canvas/src/canvas/foundation/index.ts', import.meta.url))
const canvasEngineEntry = fileURLToPath(new URL('../canvas/src/canvas/engine/index.ts', import.meta.url))
const canvasSvgDrawingPrimitivesEntry = fileURLToPath(new URL('../canvas/src/canvas/renderer/svg/CanvasSvgDrawingPrimitives.ts', import.meta.url))
const slideEditAffordanceEntry = fileURLToPath(new URL('../canvas/packages/slide-edit-affordance/src/index.ts', import.meta.url))

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: [
      { find: 'canvas/app/command-palette-items', replacement: canvasCommandPaletteItemsEntry },
      { find: 'canvas/app/eraser-hit-testing', replacement: canvasEraserHitTestingEntry },
      { find: 'canvas/app/pointer-drawing', replacement: canvasPointerDrawingEntry },
      { find: 'canvas/app/image-import', replacement: canvasImageImportEntry },
      { find: 'canvas/app/inline-edit-dom', replacement: canvasInlineEditDomEntry },
      { find: 'canvas/app/keyboard-command-dispatch', replacement: canvasKeyboardCommandDispatchEntry },
      { find: 'canvas/app/keyboard-command-shortcuts', replacement: canvasKeyboardCommandShortcutsEntry },
      { find: 'canvas/app/keyboard-nudge-shortcuts', replacement: canvasKeyboardNudgeShortcutsEntry },
      { find: 'canvas/app/keyboard-tool-shortcuts', replacement: canvasKeyboardToolShortcutsEntry },
      { find: 'canvas/app/keyboard-viewport-dispatch', replacement: canvasKeyboardViewportDispatchEntry },
      { find: 'canvas/app/keyboard-viewport-shortcuts', replacement: canvasKeyboardViewportShortcutsEntry },
      { find: 'canvas/app/menu-roving-focus', replacement: canvasMenuRovingFocusEntry },
      { find: 'canvas/app/media-import', replacement: canvasMediaImportEntry },
      { find: 'canvas/app/minimap-model', replacement: canvasMinimapModelEntry },
      { find: 'canvas/app/modal-focus-lifecycle', replacement: canvasModalFocusLifecycleEntry },
      { find: 'canvas/app/paste-position', replacement: canvasPastePositionEntry },
      { find: 'canvas/app/pointer-click-memory', replacement: canvasPointerClickMemoryEntry },
      { find: 'canvas/app/pointer-geometry', replacement: canvasPointerGeometryEntry },
      { find: 'canvas/app/pointer-laser', replacement: canvasPointerLaserEntry },
      { find: 'canvas/app/pointer-pan-interaction', replacement: canvasPointerPanInteractionEntry },
      { find: 'canvas/app/radio-group', replacement: canvasRadioGroupEntry },
      { find: 'canvas/app/rich-clipboard', replacement: canvasRichClipboardEntry },
      { find: 'canvas/app/tabs-roving-focus', replacement: canvasTabsRovingFocusEntry },
      { find: 'canvas/app/stage-element', replacement: canvasStageElementEntry },
      { find: 'canvas/app/table-import', replacement: canvasTableImportEntry },
      { find: 'canvas/app/text-paste-import', replacement: canvasTextPasteImportEntry },
      { find: 'canvas/app/toolbar-roving-focus', replacement: canvasToolbarRovingFocusEntry },
      { find: 'canvas/app/viewport-controls', replacement: canvasViewportControlsEntry },
      { find: 'canvas/core', replacement: canvasCoreEntry },
      { find: 'canvas/foundation', replacement: canvasFoundationEntry },
      { find: 'canvas/engine', replacement: canvasEngineEntry },
      { find: 'canvas/renderer/svg-drawing-primitives', replacement: canvasSvgDrawingPrimitivesEntry },
      { find: '@interactive-os/slide-edit-affordance', replacement: slideEditAffordanceEntry },
    ],
    dedupe: ['react', 'react-dom', 'zod'],
  },
})
