import { useState } from 'react'
import type { JSONDocument } from 'zod-crud'
import type { RetouchDeck, RetouchSlide } from './retouchModel'
import type { CanvasView } from './retouchViewState'

type UseRetouchPresentationCommandsParams = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  clearTransientState: () => void
  commitActiveTextEdit: () => RetouchDeck
  doc: JSONDocument<RetouchDeck>
  setActiveSlideId: (slideId: string) => void
  setCanvasView: (canvasView: CanvasView) => void
}

export function useRetouchPresentationCommands({
  activeSlide,
  activeSlideIndex,
  clearTransientState,
  commitActiveTextEdit,
  doc,
  setActiveSlideId,
  setCanvasView,
}: UseRetouchPresentationCommandsParams) {
  const [presenting, setPresenting] = useState(false)

  function startPresentation() {
    commitActiveTextEdit()
    setCanvasView('slide')
    doc.selection?.empty()
    clearTransientState()
    setPresenting(true)
  }

  function closePresentation() {
    setPresenting(false)
  }

  function navigatePresentation(direction: -1 | 1) {
    const nextIndex = Math.min(
      doc.value.slides.length - 1,
      Math.max(0, activeSlideIndex + direction),
    )
    const nextSlide = doc.value.slides[nextIndex]

    if (nextSlide === undefined || nextSlide.id === activeSlide.id) {
      return
    }

    setActiveSlideId(nextSlide.id)
    doc.selection?.empty()
    clearTransientState()
  }

  return {
    closePresentation,
    navigatePresentation,
    presenting,
    startPresentation,
  }
}
