import { Copy, Download, SlidersHorizontal } from 'lucide-react'
import {
  useEffect,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react'
import {
  createPPTCanvasTabsDescriptor,
  focusPPTCanvasElementBySelectorOnNextFrame,
  getPPTCanvasTabsKeyboardIntent,
  PPT_TABS_ROVING_FOCUS_MODEL,
  type PPTCanvasTabsDescriptor,
} from '../../pptCanvasAppAffordanceAdapter'
import { Button, IconButton, Tab, TabList, ToolbarGroup } from '../core'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import { PPTLayerPane } from './PPTLayerPane'
import { PPTSlideInspectorPanel } from './PPTSlideInspectorPanel'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTInspectorTabId = 'selection' | 'slide'
type PPTInspectorShellProps = PPTInspectorProps & {
  selectionPanel: ReactNode
}

const PPT_INSPECTOR_TABS = [
  {
    id: 'slide',
    label: 'Slide',
    panelId: 'ppt-inspector-panel-slide',
    tabId: 'ppt-inspector-tab-slide',
  },
  {
    id: 'selection',
    label: 'Selection',
    panelId: 'ppt-inspector-panel-selection',
    tabId: 'ppt-inspector-tab-selection',
  },
] as const satisfies readonly {
  id: PPTInspectorTabId
  label: string
  panelId: string
  tabId: string
}[]

export function PPTInspectorShell({
  model,
  onAction,
  selectionPanel,
}: PPTInspectorShellProps) {
  const {
    exportCode,
    hidden,
    inspectorSurface,
    selectedElement,
    slideMetadataDescriptor,
  } = model
  const { onCopyHTML, onDownloadHTML } =
    createPPTInspectorActionDispatcher(onAction)
  const hasSelectedElement = selectedElement !== null
  const [activeInspectorTabId, setActiveInspectorTabId] =
    useState<PPTInspectorTabId>(hasSelectedElement ? 'selection' : 'slide')
  const [advancedInspectorOpen, setAdvancedInspectorOpen] = useState(false)
  const previousInspectorSelectionStateRef = useRef(hasSelectedElement)
  const inspectorTabsDescriptor = createPPTCanvasTabsDescriptor({
    activation: 'automatic',
    activeId: activeInspectorTabId,
    tabs: PPT_INSPECTOR_TABS,
  })
  const slideInspectorPanelAttributes = getPPTInspectorPanelAttributes(
    inspectorTabsDescriptor,
    'slide',
  )
  const selectionInspectorPanelAttributes = getPPTInspectorPanelAttributes(
    inspectorTabsDescriptor,
    'selection',
  )

  useEffect(() => {
    if (previousInspectorSelectionStateRef.current === hasSelectedElement) {
      return
    }

    previousInspectorSelectionStateRef.current = hasSelectedElement
    setActiveInspectorTabId(hasSelectedElement ? 'selection' : 'slide')
  }, [hasSelectedElement])

  function focusPPTInspectorTab(tabId: PPTInspectorTabId) {
    focusPPTCanvasElementBySelectorOnNextFrame<HTMLButtonElement>({
      match: ({ element }) =>
        element.getAttribute('data-ppt-inspector-tab') === tabId,
      root: document,
      selector: '[data-ppt-inspector-tab]',
    })
  }

  function handlePPTInspectorTabKeyDown(
    tabId: PPTInspectorTabId,
    event: ReactKeyboardEvent<HTMLButtonElement>,
  ) {
    if (
      event.target !== event.currentTarget ||
      event.ctrlKey ||
      event.metaKey ||
      (event.altKey && event.key !== 'ArrowDown' && event.key !== 'ArrowUp')
    ) {
      return
    }

    const intent = getPPTCanvasTabsKeyboardIntent({
      activation: inspectorTabsDescriptor.activation,
      currentId: tabId,
      key: event.key,
      tabs: PPT_INSPECTOR_TABS,
    })

    if (intent.kind === 'none') {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    if (intent.activate) {
      setActiveInspectorTabId(intent.id)
    }
    focusPPTInspectorTab(intent.id)
  }

  return (
    <aside
      aria-label="Inspector"
      className="ppt-inspector"
      data-ppt-inspector-active-tab={activeInspectorTabId}
      data-ppt-inspector-advanced-open={
        advancedInspectorOpen ? 'true' : 'false'
      }
      hidden={hidden}
    >
      <div className="ppt-inspector-head">
        <TabList
          aria-label="Inspector panels"
          className="ppt-inspector-tabs"
          data-ppt-inspector-tabs
          data-ppt-inspector-tabs-activation={inspectorTabsDescriptor.activation}
          data-ppt-inspector-tabs-keyboard={inspectorTabsDescriptor.keyboardModel}
          data-ppt-inspector-tabs-model={PPT_TABS_ROVING_FOCUS_MODEL}
        >
          {inspectorTabsDescriptor.tabs.map((tab) => (
            <Tab
              {...tab.attributes}
              className="ppt-inspector-tab"
              data-ppt-inspector-tab={tab.id}
              data-ppt-inspector-tab-active={tab.isActive ? 'true' : 'false'}
              key={tab.id}
              onClick={() => setActiveInspectorTabId(tab.id)}
              onKeyDown={(event) =>
                handlePPTInspectorTabKeyDown(tab.id, event)}
            >
              {tab.label}
            </Tab>
          ))}
        </TabList>
        <IconButton
          aria-pressed={advancedInspectorOpen}
          className="ppt-inspector-advanced-toggle"
          data-ppt-inspector-advanced-toggle
          label={
            advancedInspectorOpen
              ? 'Hide advanced properties'
              : 'Show advanced properties'
          }
          onClick={() => setAdvancedInspectorOpen((current) => !current)}
        >
          <SlidersHorizontal size={15} />
        </IconButton>
      </div>

      <section
        {...slideInspectorPanelAttributes}
        className="ppt-panel-section"
        data-ppt-inspector-surface={inspectorSurface}
        data-ppt-inspector-tabpanel="slide"
        data-ppt-inspector-tabpanel-active={
          activeInspectorTabId === 'slide' ? 'true' : 'false'
        }
        data-ppt-slide-inspector-priority={
          inspectorSurface === 'slide-metadata-inspector'
            ? 'active'
            : 'secondary'
        }
        data-ppt-slide-metadata-active-index={
          slideMetadataDescriptor.activeSlide.index ?? ''
        }
        data-ppt-slide-metadata-command-slot="command-effect"
        data-ppt-slide-metadata-field-count={
          slideMetadataDescriptor.fields.length
        }
        data-ppt-slide-metadata-inspector
        data-ppt-slide-metadata-slide-count={
          slideMetadataDescriptor.activeSlide.slideCount
        }
        data-ppt-slide-metadata-slide-id={
          slideMetadataDescriptor.activeSlide.slideId
        }
        data-ppt-slide-metadata-surface={slideMetadataDescriptor.surface}
      >
        <PPTSlideInspectorPanel model={model} onAction={onAction} />
      </section>

      <section
        {...selectionInspectorPanelAttributes}
        className="ppt-panel-section"
        data-ppt-inspector-surface={inspectorSurface}
        data-ppt-inspector-tabpanel="selection"
        data-ppt-inspector-tabpanel-active={
          activeInspectorTabId === 'selection' ? 'true' : 'false'
        }
        data-ppt-object-inspector
        data-ppt-object-inspector-active={selectedElement ? 'true' : 'false'}
        data-ppt-object-inspector-priority={
          inspectorSurface === 'object-selection-inspector' ? '0' : '1'
        }
      >
        {selectionPanel}
      </section>

      <PPTLayerPane model={model} onAction={onAction} />

      <div className="ppt-panel-header ppt-inspector-advanced-only">
        <h2>Export</h2>
      </div>
      <section className="ppt-panel-section ppt-inspector-advanced-only">
        <ToolbarGroup>
          <Button onClick={onCopyHTML}>
            <Copy size={16} /> Copy
          </Button>
          <Button onClick={onDownloadHTML}>
            <Download size={16} /> Download
          </Button>
        </ToolbarGroup>
        <textarea className="ppt-export-code" readOnly value={exportCode} />
      </section>
    </aside>
  )
}

function getPPTInspectorPanelAttributes(
  descriptor: PPTCanvasTabsDescriptor<PPTInspectorTabId>,
  tabId: PPTInspectorTabId,
) {
  const panel = descriptor.panels.find((item) => item.id === tabId)

  if (!panel) {
    throw new Error(`Missing PPT inspector panel for tab ${tabId}`)
  }

  return panel.attributes
}
