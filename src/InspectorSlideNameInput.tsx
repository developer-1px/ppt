import { useRef, useState } from 'react'

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
