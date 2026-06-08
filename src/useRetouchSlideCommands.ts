import { useState, type RefObject } from 'react'
import type { JSONDocument, JSONPatchOperation } from 'zod-crud'
import type { RetouchCollection } from './retouchCollection'
import type { RetouchDeck, RetouchSlide } from './retouchModel'
import {
  addSlidePatch,
  createBlankSlide,
  duplicateSlide,
  setSlideAccentPatch,
  setSlideNamePatch,
} from './slideDeckOperations'
import type { CanvasView } from './retouchViewState'

type UseRetouchSlideCommandsParams = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  clearTransientState: () => void
  commitActiveTextEdit: () => RetouchDeck
  commitRetouchPatch: (
    patch: readonly JSONPatchOperation[],
    label: string,
  ) => void
  doc: JSONDocument<RetouchDeck>
  retouchCollection: RetouchCollection
  setActiveSlideId: (slideId: string) => void
  setCanvasView: (canvasView: CanvasView) => void
  stageRef: RefObject<HTMLDivElement | null>
}

export function useRetouchSlideCommands({
  activeSlide,
  activeSlideIndex,
  clearTransientState,
  commitActiveTextEdit,
  commitRetouchPatch,
  doc,
  retouchCollection,
  setActiveSlideId,
  setCanvasView,
  stageRef,
}: UseRetouchSlideCommandsParams) {
  const [notesBySlideId, setNotesBySlideId] = useState<Record<string, string>>({})

  function activateSlide(
    slideId: string,
    { resetStageScroll = true }: { resetStageScroll?: boolean } = {},
  ) {
    setActiveSlideId(slideId)
    setCanvasView('slide')
    doc.selection?.empty()
    if (resetStageScroll) {
      stageRef.current?.scrollTo({ left: 0, top: 0 })
    }
    clearTransientState()
  }

  function selectSlide(slideId: string) {
    commitActiveTextEdit()
    activateSlide(slideId)
  }

  function addSlide() {
    commitActiveTextEdit()
    const nextSlide = createBlankSlide(doc.value.slides)
    const insertIndex = activeSlideIndex + 1

    commitRetouchPatch(addSlidePatch(nextSlide, insertIndex), 'add slide')
    activateSlide(nextSlide.id)
  }

  function copySlide() {
    commitActiveTextEdit()
    const nextSlide = duplicateSlide(activeSlide, doc.value.slides)
    const insertIndex = activeSlideIndex + 1

    commitRetouchPatch(addSlidePatch(nextSlide, insertIndex), 'duplicate slide')
    setNotesBySlideId((current) => ({
      ...current,
      [nextSlide.id]: current[activeSlide.id] ?? '',
    }))
    activateSlide(nextSlide.id)
  }

  function deleteSlide() {
    if (doc.value.slides.length <= 1) {
      return
    }

    commitActiveTextEdit()
    const nextSlide =
      doc.value.slides[activeSlideIndex + 1] ??
      doc.value.slides[activeSlideIndex - 1] ??
      doc.value.slides[0]

    const deleted = retouchCollection.deleteSlide(activeSlideIndex)
    if (!deleted.ok) {
      return
    }

    setNotesBySlideId((current) => {
      const next = { ...current }
      delete next[activeSlide.id]
      return next
    })
    activateSlide(nextSlide.id)
  }

  function moveSlide(direction: -1 | 1) {
    const nextIndex = activeSlideIndex + direction

    if (nextIndex < 0 || nextIndex >= doc.value.slides.length) {
      return
    }

    commitActiveTextEdit()
    const moved = retouchCollection.moveSlide(activeSlideIndex, direction)
    if (!moved.ok) {
      return
    }

    activateSlide(activeSlide.id, { resetStageScroll: false })
  }

  function changeSlideName(name: string) {
    const nextName = name.trim() || 'Untitled'

    if (nextName === activeSlide.name) {
      return
    }

    commitActiveTextEdit()
    commitRetouchPatch(setSlideNamePatch(activeSlideIndex, nextName), 'rename slide')
    clearTransientState()
  }

  function changeSlideAccent(accent: string) {
    if (accent === activeSlide.accent) {
      return
    }

    commitActiveTextEdit()
    commitRetouchPatch(setSlideAccentPatch(activeSlideIndex, accent), 'change slide accent')
    clearTransientState()
  }

  function changeActiveSlideNotes(notes: string) {
    setNotesBySlideId((current) => ({
      ...current,
      [activeSlide.id]: notes,
    }))
  }

  return {
    activeSlideNotes: notesBySlideId[activeSlide.id] ?? '',
    addSlide,
    changeActiveSlideNotes,
    changeSlideAccent,
    changeSlideName,
    copySlide,
    deleteSlide,
    moveSlide,
    selectSlide,
  }
}
