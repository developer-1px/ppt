import { useRef, useState } from 'react'
import { Play } from 'lucide-react'
import type { SlideBlock } from './retouchModel'
import { SLIDE_ACCENTS } from './slideDeckOperations'

type InspectorPanelProps = {
  activeSlideAccent: string
  activeSlideName: string
  canvasView: 'slide' | 'grid'
  mode: 'text' | 'layout'
  notes: string
  onNotesChange: (notes: string) => void
  onPresent: () => void
  onSlideAccentChange: (accent: string) => void
  onSlideNameChange: (name: string) => void
  selectedBlock: SlideBlock | null
}

export function InspectorPanel({
  activeSlideAccent,
  activeSlideName,
  canvasView,
  mode,
  notes,
  onNotesChange,
  onPresent,
  onSlideAccentChange,
  onSlideNameChange,
  selectedBlock,
}: InspectorPanelProps) {
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
        <h2>Selection</h2>
        {selectedBlock ? (
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
