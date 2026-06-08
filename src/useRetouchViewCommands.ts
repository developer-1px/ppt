import type { CanvasView, RetouchMode } from './retouchViewState'
import { useCanvasViewTabs } from './useCanvasViewTabs'

type UseRetouchViewCommandsParams = {
  canvasView: CanvasView
  clearSelection: () => void
  clearTransientState: () => void
  commitActiveTextEdit: () => void
  mode: RetouchMode
  setCanvasView: (canvasView: CanvasView) => void
  setMode: (mode: RetouchMode) => void
}

export function useRetouchViewCommands({
  canvasView,
  clearSelection,
  clearTransientState,
  commitActiveTextEdit,
  mode,
  setCanvasView,
  setMode,
}: UseRetouchViewCommandsParams) {
  function enterLayoutMode() {
    setCanvasView('slide')
    setMode('layout')
    clearTransientState()
  }

  function changeMode(nextMode: RetouchMode) {
    if (mode === 'text') {
      commitActiveTextEdit()
    }

    if (mode !== 'text' && nextMode === 'text') {
      clearSelection()
    }

    setMode(nextMode)
    clearTransientState()
  }

  function changeCanvasView(nextView: CanvasView) {
    if (nextView === 'grid') {
      commitActiveTextEdit()
      clearSelection()
      clearTransientState()
    }

    setCanvasView(nextView)
  }

  const canvasViewTabs = useCanvasViewTabs({
    canvasView,
    onChange: changeCanvasView,
  })

  return {
    ...canvasViewTabs,
    changeMode,
    enterLayoutMode,
  }
}
