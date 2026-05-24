import { useRef, useState } from 'react'
import { Copy, Play, Trash2 } from 'lucide-react'
import type { Rect, SlideBlock } from './retouchModel'
import { SLIDE_ACCENTS } from './slideDeckOperations'

type RectField = keyof Rect

type InspectorPanelProps = {
  activeSlideAccent: string
  activeSlideName: string
  canvasView: 'slide' | 'grid'
  mode: 'text' | 'layout'
  notes: string
  onDeleteBlock: () => void
  onDuplicateBlock: () => void
  onNotesChange: (notes: string) => void
  onPresent: () => void
  onSelectedRectChange: (rect: Rect, changedField?: RectField) => void
  onSlideAccentChange: (accent: string) => void
  onSlideNameChange: (name: string) => void
  selectedBlock: SlideBlock | null
  selectedCount: number
  selectedRect: Rect | null
}

const GEOMETRY_FIELDS: { field: RectField; label: string }[] = [
  { field: 'x', label: 'X' },
  { field: 'y', label: 'Y' },
  { field: 'width', label: 'W' },
  { field: 'height', label: 'H' },
]

export function InspectorPanel({
  activeSlideAccent,
  activeSlideName,
  canvasView,
  mode,
  notes,
  onDeleteBlock,
  onDuplicateBlock,
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
              <div className="accent-swatches" role="radiogroup" aria-label="Slide accent">
                {SLIDE_ACCENTS.map((accent) => (
                  <button
                    aria-label={accent}
                    aria-checked={accent === activeSlideAccent}
                    className="accent-swatch"
                    key={accent}
                    onClick={() => onSlideAccentChange(accent)}
                    role="radio"
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
          <dl>
            <div>
              <dt>Blocks</dt>
              <dd>{selectedCount} selected</dd>
            </div>
          </dl>
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

function rectDraft(rect: Rect): Record<RectField, string> {
  return {
    x: String(rect.x),
    y: String(rect.y),
    width: String(rect.width),
    height: String(rect.height),
  }
}

function GeometryEditor({
  onRectChange,
  rect,
}: {
  onRectChange: (rect: Rect, changedField?: RectField) => void
  rect: Rect
}) {
  const [draft, setDraft] = useState(() => rectDraft(rect))
  const draftRef = useRef(draft)
  const skipNextCommitRef = useRef<RectField | null>(null)

  function updateDraftField(field: RectField, value: string) {
    const nextDraft = {
      ...draftRef.current,
      [field]: value,
    }

    draftRef.current = nextDraft
    setDraft(nextDraft)
  }

  function commitField(field: RectField) {
    if (skipNextCommitRef.current === field) {
      skipNextCommitRef.current = null
      return
    }

    const value = Number(draftRef.current[field])

    if (!Number.isFinite(value)) {
      updateDraftField(field, String(rect[field]))
      return
    }

    if (value === rect[field]) {
      return
    }

    onRectChange({ ...rect, [field]: value }, field)
  }

  return (
    <div aria-label="Geometry" className="geometry-editor" role="group">
      {GEOMETRY_FIELDS.map(({ field, label }) => (
        <label className="geometry-field" key={field}>
          <span>{label}</span>
          <input
            aria-label={label}
            inputMode="numeric"
            min={field === 'width' || field === 'height' ? 72 : 0}
            onBlur={() => commitField(field)}
            onChange={(event) => {
              const value =
                event.target instanceof HTMLInputElement ? event.target.value : ''

              updateDraftField(field, value)
            }}
            onFocus={(event) => event.currentTarget.select()}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.currentTarget.blur()
              }

              if (event.key === 'Escape') {
                skipNextCommitRef.current = field
                updateDraftField(field, String(rect[field]))
                event.currentTarget.blur()
              }
            }}
            step={8}
            type="number"
            value={draft[field]}
          />
        </label>
      ))}
    </div>
  )
}

function SlideNameInput({
  activeSlideName,
  onSlideNameChange,
}: {
  activeSlideName: string
  onSlideNameChange: (name: string) => void
}) {
  const [nameDraft, setNameDraft] = useState(activeSlideName)
  const skipNextNameCommitRef = useRef(false)

  function commitNameDraft() {
    if (skipNextNameCommitRef.current) {
      skipNextNameCommitRef.current = false
      return
    }

    const nextName = nameDraft.trim()

    onSlideNameChange(nextName)
    setNameDraft(nextName || 'Untitled')
  }

  return (
    <input
      aria-label="Slide name"
      className="inspector-name-input"
      onBlur={commitNameDraft}
      onChange={(event) => setNameDraft(event.currentTarget.value)}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.currentTarget.blur()
        }

        if (event.key === 'Escape') {
          skipNextNameCommitRef.current = true
          setNameDraft(activeSlideName)
          event.currentTarget.blur()
        }
      }}
      spellCheck={false}
      value={nameDraft}
    />
  )
}
