import { useLayoutEffect, useState, type RefObject } from 'react'
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  type Rect,
  type SlideBlock,
} from './retouchModel'
import { rectClose, type DraftLayout } from './layoutInteraction'
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

    const slideBox = slideRef.current.getBoundingClientRect()
    const block = slideRef.current.querySelector<HTMLElement>(
      `[data-block="${selectedBlock.id}"]`,
    )

    if (!block || slideBox.width === 0 || slideBox.height === 0) {
      setVisualSelectionRect(null)
      return
    }

    const blockBox = block.getBoundingClientRect()
    const nextRect = {
      x: ((blockBox.left - slideBox.left) / slideBox.width) * SLIDE_WIDTH,
      y: ((blockBox.top - slideBox.top) / slideBox.height) * SLIDE_HEIGHT,
      width: (blockBox.width / slideBox.width) * SLIDE_WIDTH,
      height: (blockBox.height / slideBox.height) * SLIDE_HEIGHT,
    }

    setVisualSelectionRect((currentRect) =>
      currentRect && rectClose(currentRect, nextRect) ? currentRect : nextRect,
    )
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
