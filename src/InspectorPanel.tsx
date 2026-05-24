import { Play } from 'lucide-react'
import type { SlideBlock } from './retouchModel'

type InspectorPanelProps = {
  activeSlideName: string
  canvasView: 'slide' | 'grid'
  mode: 'text' | 'layout'
  notes: string
  onNotesChange: (notes: string) => void
  onPresent: () => void
  selectedBlock: SlideBlock | null
}

export function InspectorPanel({
  activeSlideName,
  canvasView,
  mode,
  notes,
  onNotesChange,
  onPresent,
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
            <dd>{activeSlideName}</dd>
          </div>
          <div>
            <dt>View</dt>
            <dd>{canvasView === 'grid' ? 'Grid' : 'Slide'}</dd>
          </div>
          <div>
            <dt>Mode</dt>
            <dd>{mode === 'layout' ? 'Arrange' : 'Text'}</dd>
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
