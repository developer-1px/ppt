import { useCallback, useEffect, useReducer } from 'react'

export type EditorTransientSurface = 'export' | 'tools' | 'view'
export type EditorViewFeature = 'frameGuides' | 'grid' | 'minimap'

type EditorChromeState = {
  activeTransientSurface: EditorTransientSurface | null
  inspectorOpen: boolean
  view: Record<EditorViewFeature, boolean>
}

type EditorChromeAction =
  | { type: 'close-transient' }
  | { surface: EditorTransientSurface; type: 'toggle-transient' }
  | { feature: EditorViewFeature; type: 'toggle-view' }
  | { type: 'toggle-inspector' }

const INITIAL_EDITOR_CHROME_STATE: EditorChromeState = {
  activeTransientSurface: null,
  inspectorOpen: false,
  view: {
    frameGuides: false,
    grid: false,
    minimap: false,
  },
}

function reduceEditorChrome(
  state: EditorChromeState,
  action: EditorChromeAction,
): EditorChromeState {
  switch (action.type) {
    case 'close-transient':
      return state.activeTransientSurface === null
        ? state
        : { ...state, activeTransientSurface: null }
    case 'toggle-inspector':
      return { ...state, inspectorOpen: !state.inspectorOpen }
    case 'toggle-transient':
      return {
        ...state,
        activeTransientSurface:
          state.activeTransientSurface === action.surface ? null : action.surface,
      }
    case 'toggle-view':
      return {
        ...state,
        view: {
          ...state.view,
          [action.feature]: !state.view[action.feature],
        },
      }
  }
}

export function useEditorChrome() {
  const [state, dispatch] = useReducer(
    reduceEditorChrome,
    INITIAL_EDITOR_CHROME_STATE,
  )

  const closeTransient = useCallback(() => {
    dispatch({ type: 'close-transient' })
  }, [])

  const isTransientOpen = useCallback(
    (surface: EditorTransientSurface) =>
      state.activeTransientSurface === surface,
    [state.activeTransientSurface],
  )

  const isViewVisible = useCallback(
    (feature: EditorViewFeature) => state.view[feature],
    [state.view],
  )

  const toggleInspector = useCallback(() => {
    dispatch({ type: 'toggle-inspector' })
  }, [])

  const toggleTransient = useCallback((surface: EditorTransientSurface) => {
    dispatch({ surface, type: 'toggle-transient' })
  }, [])

  const toggleView = useCallback((feature: EditorViewFeature) => {
    dispatch({ feature, type: 'toggle-view' })
  }, [])

  useEffect(() => {
    if (state.activeTransientSurface === null) {
      return
    }

    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        closeTransient()
      }
    }

    window.addEventListener('keydown', closeOnEscape)
    return () => window.removeEventListener('keydown', closeOnEscape)
  }, [closeTransient, state.activeTransientSurface])

  return {
    activeTransientSurface: state.activeTransientSurface,
    closeTransient,
    inspectorOpen: state.inspectorOpen,
    isTransientOpen,
    isViewVisible,
    toggleInspector,
    toggleTransient,
    toggleView,
  }
}
