import {
  useRef,
  useState,
  type ChangeEvent as ReactChangeEvent,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'

export function SlideNameInput({
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

  function handleNameDraftChange(event: ReactChangeEvent<HTMLInputElement>) {
    setNameDraft(event.currentTarget.value)
  }

  function cancelNameDraft(input: HTMLInputElement) {
    skipNextNameCommitRef.current = true
    setNameDraft(activeSlideName)
    input.blur()
  }

  function handleNameDraftKeyDown(event: ReactKeyboardEvent<HTMLInputElement>) {
    if (event.key === 'Enter') {
      event.currentTarget.blur()
      return
    }

    if (event.key === 'Escape') {
      cancelNameDraft(event.currentTarget)
    }
  }

  return (
    <input
      aria-label="Slide name"
      className="inspector-name-input"
      onBlur={commitNameDraft}
      onChange={handleNameDraftChange}
      onKeyDown={handleNameDraftKeyDown}
      spellCheck={false}
      value={nameDraft}
    />
  )
}
