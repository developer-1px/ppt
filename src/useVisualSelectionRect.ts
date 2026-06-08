import { useLayoutEffect, useState, type RefObject } from 'react'
import {
  type Rect,
  type SlideBlock,
} from './retouchModel'
import { rectClose, type DraftLayout } from './layoutInteraction'
import { readSlideBlockRect } from './retouchSlideDom'
import type { RetouchMode } from './retouchViewState'

export function useVisualSelectionRect({
  activeSlideId,
  deckValue,
  draftLayout,
  mode,
  selectedBlock,
  selectedPointerCount,
  slideRef,
}: {
  activeSlideId: string
  deckValue: unknown
  draftLayout: DraftLayout | null
  mode: RetouchMode
  selectedBlock: SlideBlock | null
  selectedPointerCount: number
  slideRef: RefObject<HTMLDivElement | null>
}) {
  const [visualSelectionRect, setVisualSelectionRect] = useState<Rect | null>(null)

  useLayoutEffect(() => {
    if (
      mode !== 'layout' ||
      selectedPointerCount !== 1 ||
      !selectedBlock ||
      !slideRef.current
    ) {
      setVisualSelectionRect(null)
      return
    }

    const nextRect = readSlideBlockRect(slideRef.current, selectedBlock.id)

    setVisualSelectionRect((currentRect) => {
      if (!nextRect) {
        return null
      }

      return currentRect && rectClose(currentRect, nextRect)
        ? currentRect
        : nextRect
    })
  }, [
    activeSlideId,
    deckValue,
    draftLayout,
    mode,
    selectedBlock,
    selectedPointerCount,
    slideRef,
  ])

  return visualSelectionRect
}
