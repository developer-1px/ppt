import { useManagedTabsPattern } from './apgPatternAdapter'
import type { PatternElementProps } from './apgPatternTypes'
import type { CanvasView } from './retouchViewState'

type CanvasViewTabProps = Record<
  CanvasView,
  PatternElementProps
>

type UseCanvasViewTabsInput = {
  canvasView: CanvasView
  onChange: (nextView: CanvasView) => void
}

type UseCanvasViewTabsResult = {
  canvasViewPanelProps: PatternElementProps
  canvasViewTablistProps: PatternElementProps
  canvasViewTabProps: CanvasViewTabProps
}

const CANVAS_VIEW_TABS = [
  { label: 'Slide', panelKey: 'slide-panel', tabKey: 'slide', value: 'slide' },
  { label: 'Grid', panelKey: 'grid-panel', tabKey: 'grid', value: 'grid' },
] as const satisfies readonly {
  label: string
  panelKey: string
  tabKey: string
  value: CanvasView
}[]

export function useCanvasViewTabs({
  canvasView,
  onChange,
}: UseCanvasViewTabsInput): UseCanvasViewTabsResult {
  const tabs = useManagedTabsPattern<CanvasView>({
    activeValue: canvasView,
    elementIdPrefix: 'canvas-view-',
    label: 'Canvas view',
    onSelect: onChange,
    tabs: CANVAS_VIEW_TABS,
  })

  return {
    canvasViewPanelProps: tabs.panelProps,
    canvasViewTablistProps: tabs.tablistProps,
    canvasViewTabProps: {
      grid: tabs.tabPropsByValue.grid,
      slide: tabs.tabPropsByValue.slide,
    },
  }
}
