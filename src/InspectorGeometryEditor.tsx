import {
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'
import type { Rect } from './retouchModel'
import type { RectField } from './inspectorGeometry'

const GEOMETRY_FIELDS: { field: RectField; label: string }[] = [
  { field: 'x', label: 'X' },
  { field: 'y', label: 'Y' },
  { field: 'width', label: 'W' },
  { field: 'height', label: 'H' },
]

export function GeometryEditor({
  onRectChange,
  rect,
}: {
  onRectChange: (rect: Rect, changedField?: RectField) => void
  rect: Rect
}) {
  const [draft, setDraft] = useState<Record<RectField, string>>(() => ({
    x: String(rect.x),
    y: String(rect.y),
    width: String(rect.width),
    height: String(rect.height),
  }))
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

  function cancelFieldDraft(field: RectField, input: HTMLInputElement) {
    skipNextCommitRef.current = field
    updateDraftField(field, String(rect[field]))
    input.blur()
  }

  function handleFieldKeyDown(
    field: RectField,
    event: ReactKeyboardEvent<HTMLInputElement>,
  ) {
    if (event.key === 'Enter') {
      event.currentTarget.blur()
      return
    }

    if (event.key === 'Escape') {
      cancelFieldDraft(field, event.currentTarget)
    }
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
            onChange={(event) => updateDraftField(field, event.currentTarget.value)}
            onFocus={(event) => event.currentTarget.select()}
            onKeyDown={(event) => handleFieldKeyDown(field, event)}
            step={8}
            type="number"
            value={draft[field]}
          />
        </label>
      ))}
    </div>
  )
}
