import {
  Copy,
  Play,
  Trash2,
} from 'lucide-react'
import type { Rect, SlideBlock } from './retouchModel'
import type { RectField } from './inspectorGeometry'
import type { CanvasView, RetouchMode } from './retouchViewState'
import type {
  AlignSelectionAction,
  DistributeSelectionAction,
} from './selectionAlignment'
import type { LayerOrderAction } from './selectionLayerOrder'
import { SLIDE_ACCENTS } from './slideDeckOperations'
import { useManagedRadioGroupPattern } from './apgPatternAdapter'
import { GeometryEditor } from './InspectorGeometryEditor'
import {
  AlignmentTools,
  DistributionTools,
  LayerOrderTools,
} from './InspectorSelectionTools'
import { SlideNameInput } from './InspectorSlideNameInput'
import './InspectorPanel.css'

type InspectorPanelProps = {
  activeSlideAccent: string
  activeSlideName: string
  canvasView: CanvasView
  mode: RetouchMode
  notes: string
  onAlignSelection: (action: AlignSelectionAction) => void
  onDeleteBlock: () => void
  onDistributeSelection: (action: DistributeSelectionAction) => void
  onDuplicateBlock: () => void
  onLayerOrderChange: (action: LayerOrderAction) => void
  onNotesChange: (notes: string) => void
  onPresent: () => void
  onSelectedRectChange: (rect: Rect, changedField?: RectField) => void
  onSlideAccentChange: (accent: string) => void
  onSlideNameChange: (name: string) => void
  selectedBlock: SlideBlock | null
  selectedCount: number
  selectedRect: Rect | null
}

const SLIDE_ACCENT_ITEMS = Object.fromEntries(
  SLIDE_ACCENTS.map((accent) => [accent, { label: accent }]),
) as Record<string, { label: string }>

export function InspectorPanel({
  activeSlideAccent,
  activeSlideName,
  canvasView,
  mode,
  notes,
  onAlignSelection,
  onDeleteBlock,
  onDistributeSelection,
  onDuplicateBlock,
  onLayerOrderChange,
  onNotesChange,
  onPresent,
  onSelectedRectChange,
  onSlideAccentChange,
  onSlideNameChange,
  selectedBlock,
  selectedCount,
  selectedRect,
}: InspectorPanelProps) {
  const hasSelection = selectedCount > 0
  const hasMultiSelection = selectedCount > 1
  const slideAccentRadio = useManagedRadioGroupPattern<string>({
    elementIdPrefix: 'slide-accent-',
    items: SLIDE_ACCENT_ITEMS,
    label: 'Slide accent',
    onSelect: onSlideAccentChange,
    rootKeys: SLIDE_ACCENTS,
    selectedKey: activeSlideAccent,
  })

  return (
    <aside className="inspector-panel" aria-label="Slide details">
      <section className="inspector-section">
        <div className="inspector-heading">
          <h2>Slide</h2>
          <button aria-label="Present" onClick={onPresent} title="Present" type="button">
            <Play aria-hidden="true" size={15} strokeWidth={2.2} />
          </button>
        </div>
        <dl>
          <div>
            <dt>Name</dt>
            <dd>
              <SlideNameInput
                key={activeSlideName}
                activeSlideName={activeSlideName}
                onSlideNameChange={onSlideNameChange}
              />
            </dd>
          </div>
          <div>
            <dt>View</dt>
            <dd>{canvasView === 'grid' ? 'Grid' : 'Slide'}</dd>
          </div>
          <div>
            <dt>Mode</dt>
            <dd>{mode === 'layout' ? 'Arrange' : 'Text'}</dd>
          </div>
          <div>
            <dt>Accent</dt>
            <dd>
              <div {...slideAccentRadio.rootProps} className="accent-swatches">
                {SLIDE_ACCENTS.map((accent) => (
                  <button
                    {...slideAccentRadio.itemProps[accent]}
                    aria-label={accent}
                    className="accent-swatch"
                    key={accent}
                    style={{ background: accent }}
                    type="button"
                  />
                ))}
              </div>
            </dd>
          </div>
        </dl>
      </section>

      <section className="inspector-section">
        <div className="inspector-heading">
          <h2>Selection</h2>
          {hasSelection ? (
            <div className="selection-actions">
              <button
                aria-label={
                  hasMultiSelection ? 'Duplicate selection' : 'Duplicate block'
                }
                onClick={onDuplicateBlock}
                title={hasMultiSelection ? 'Duplicate selection' : 'Duplicate block'}
                type="button"
              >
                <Copy aria-hidden="true" size={15} strokeWidth={2.2} />
              </button>
              <button
                aria-label={hasMultiSelection ? 'Delete selection' : 'Delete block'}
                onClick={onDeleteBlock}
                title={hasMultiSelection ? 'Delete selection' : 'Delete block'}
                type="button"
              >
                <Trash2 aria-hidden="true" size={15} strokeWidth={2.2} />
              </button>
            </div>
          ) : null}
        </div>
        {hasMultiSelection ? (
          <>
            <dl>
              <div>
                <dt>Blocks</dt>
                <dd>{selectedCount} selected</dd>
              </div>
            </dl>
            <AlignmentTools onAlignSelection={onAlignSelection} />
            <DistributionTools
              disabled={selectedCount < 3}
              onDistributeSelection={onDistributeSelection}
            />
            <LayerOrderTools onLayerOrderChange={onLayerOrderChange} />
          </>
        ) : selectedBlock ? (
          <>
            <dl>
              <div>
                <dt>Block</dt>
                <dd>{selectedBlock.id}</dd>
              </div>
              <div>
                <dt>Role</dt>
                <dd>{selectedBlock.role}</dd>
              </div>
            </dl>
            <AlignmentTools onAlignSelection={onAlignSelection} />
            <LayerOrderTools onLayerOrderChange={onLayerOrderChange} />
            {selectedRect ? (
              <GeometryEditor
                key={`${selectedBlock.id}:${selectedRect.x}:${selectedRect.y}:${selectedRect.width}:${selectedRect.height}`}
                onRectChange={onSelectedRectChange}
                rect={selectedRect}
              />
            ) : null}
          </>
        ) : (
          <p className="inspector-empty">None</p>
        )}
      </section>

      <section className="inspector-section inspector-notes">
        <label htmlFor="slide-notes">Notes</label>
        <textarea
          id="slide-notes"
          onChange={(event) => onNotesChange(event.currentTarget.value)}
          spellCheck={false}
          value={notes}
        />
      </section>
    </aside>
  )
}
