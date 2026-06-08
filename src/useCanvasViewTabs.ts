import {
  type ButtonHTMLAttributes,
  type HTMLAttributes,
  useMemo,
} from 'react'
import {
  reducePatternData,
  tabsDefinition,
  useTabsPattern,
  type PatternData,
  type PatternEvent,
} from '@interactive-os/aria/react'
import {
  patternButtonProps,
  patternDivProps,
} from './apgPatternAdapter'

export type CanvasView = 'slide' | 'grid'

type CanvasViewTabProps = Record<
  CanvasView,
  ButtonHTMLAttributes<HTMLButtonElement>
>

type UseCanvasViewTabsInput = {
  canvasView: CanvasView
  onChange: (nextView: CanvasView) => void
}

type UseCanvasViewTabsResult = {
  canvasViewPanelProps: HTMLAttributes<HTMLDivElement>
  canvasViewTablistProps: HTMLAttributes<HTMLDivElement>
  canvasViewTabProps: CanvasViewTabProps
}

const CANVAS_VIEW_TABS = [
  { label: 'Slide', panelKey: 'slide-panel', tabKey: 'slide', view: 'slide' },
  { label: 'Grid', panelKey: 'grid-panel', tabKey: 'grid', view: 'grid' },
] as const satisfies readonly {
  label: string
  panelKey: string
  tabKey: string
  view: CanvasView
}[]

export function useCanvasViewTabs({
  canvasView,
  onChange,
}: UseCanvasViewTabsInput): UseCanvasViewTabsResult {
  const tabData = useMemo<PatternData>(() => {
    const activeTab = canvasViewTab(canvasView)

    return {
      items: Object.fromEntries(
        CANVAS_VIEW_TABS.flatMap((tab) => [
          [tab.tabKey, { label: tab.label }],
          [tab.panelKey, { label: `${tab.label} view` }],
        ]),
      ),
      relations: {
        controlsByKey: Object.fromEntries(
          CANVAS_VIEW_TABS.map((tab) => [tab.tabKey, [tab.panelKey]]),
        ),
        ownerByKey: Object.fromEntries(
          CANVAS_VIEW_TABS.map((tab) => [tab.panelKey, tab.tabKey]),
        ),
        rootKeys: CANVAS_VIEW_TABS.map((tab) => tab.tabKey),
      },
      refs: { label: 'Canvas view' },
      state: {
        activeKey: activeTab.tabKey,
        selectedKeys: [activeTab.tabKey],
      },
    }
  }, [canvasView])
  const tabs = useTabsPattern(
    tabData,
    (event: PatternEvent) => {
      const nextView = canvasViewFromTabsEvent(tabData, event)

      if (nextView) {
        onChange(nextView)
      }
    },
    {
      activationMode: 'automatic',
      elementIdPrefix: 'canvas-view-',
      orientation: 'horizontal',
    },
  )

  return {
    canvasViewPanelProps: patternDivProps(
      tabs.getTabPanelProps(canvasViewTab(canvasView).panelKey),
    ),
    canvasViewTablistProps: patternDivProps(tabs.getTablistProps()),
    canvasViewTabProps: {
      grid: patternButtonProps(tabs.getTabProps(canvasViewTab('grid').tabKey)),
      slide: patternButtonProps(tabs.getTabProps(canvasViewTab('slide').tabKey)),
    },
  }
}

function canvasViewTab(view: CanvasView) {
  return CANVAS_VIEW_TABS.find((tab) => tab.view === view) ?? CANVAS_VIEW_TABS[0]
}

function canvasViewFromTabKey(
  tabKey: string | null | undefined,
): CanvasView | null {
  return CANVAS_VIEW_TABS.find((tab) => tab.tabKey === tabKey)?.view ?? null
}

function canvasViewFromTabsEvent(
  data: PatternData,
  event: PatternEvent,
): CanvasView | null {
  if (event.type === 'select') {
    return canvasViewFromTabKey(event.keys[0])
  }

  if (event.type === 'navigate') {
    const nextData = reducePatternData(tabsDefinition, data, event)

    return canvasViewFromTabKey(nextData.state?.activeKey)
  }

  return null
}
