import {
  useCallback,
  useEffect,
  useState,
  type MutableRefObject,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import { pointDistance } from 'canvas/core'
import type { Pointer } from 'zod-crud'
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  blockLocationFromPointer,
  resizeRect,
  setArrangePatch,
  type Rect,
  type RetouchDeck,
  type ResizeHandle,
  type SlideBlock,
} from './retouchModel'
import {
  calculateMoveInteractionState,
  clamp,
  draftRectsEqual,
  getCurrentRect,
  guidesForInteraction,
  hasSelectionModifier,
  selectionSnapForPointers,
  type DraftLayout,
  type DraftLayoutRect,
  type Interaction,
  type Point,
  type SnapGuides,
} from './layoutInteraction'
import type { RetouchSurfaceCommitPatch } from './retouchSurfaceContract'
import type { RetouchMode } from './retouchViewState'

const DRAG_THRESHOLD = 8

const hasMeaningfulPointerDelta = (start: Point, next: Point) =>
  pointDistance(start, next) >= DRAG_THRESHOLD

export function useRetouchLayoutInteraction({
  activeSlideBlocks,
  activeSlideId,
  activeSlideIndex,
  commitPatch,
  deckValue,
  mode,
  selectedBlock,
  selectedPointer,
  selectedPointerSet,
  selectedPointers,
  selectBlock,
  slideRef,
  suppressBlockClickRef,
  suppressStageClickRef,
}: {
  activeSlideBlocks: SlideBlock[]
  activeSlideId: string
  activeSlideIndex: number
  commitPatch: RetouchSurfaceCommitPatch
  deckValue: RetouchDeck
  mode: RetouchMode
  selectedBlock: SlideBlock | null
  selectedPointer: Pointer | null
  selectedPointerSet: Set<Pointer>
  selectedPointers: Pointer[]
  selectBlock: (pointer: Pointer) => void
  slideRef: RefObject<HTMLDivElement | null>
  suppressBlockClickRef: MutableRefObject<boolean>
  suppressStageClickRef: MutableRefObject<boolean>
}) {
  const [interaction, setInteraction] = useState<Interaction | null>(null)
  const [draftLayout, setDraftLayout] = useState<DraftLayout | null>(null)
  const [snapGuides, setSnapGuides] = useState<SnapGuides>({
    x: null,
    y: null,
  })

  const clearLayoutInteraction = useCallback(() => {
    setInteraction(null)
    setDraftLayout(null)
    setSnapGuides({ x: null, y: null })
  }, [])

  const readSlidePoint = useCallback((event: Pick<PointerEvent, 'clientX' | 'clientY'>) => {
    const rect = slideRef.current?.getBoundingClientRect()

    if (!rect) {
      return null
    }

    return {
      x: clamp(((event.clientX - rect.left) / rect.width) * SLIDE_WIDTH, 0, SLIDE_WIDTH),
      y: clamp(((event.clientY - rect.top) / rect.height) * SLIDE_HEIGHT, 0, SLIDE_HEIGHT),
    }
  }, [slideRef])

  const readBlockVisualRect = useCallback((blockId: string, fallback: Rect) => {
    const slideElement = slideRef.current
    const slideRect = slideElement?.getBoundingClientRect()
    const blockElement = Array.from(
      slideElement?.querySelectorAll<HTMLElement>('[data-block]') ?? [],
    ).find((element) => element.dataset.block === blockId)
    const blockRect = blockElement?.getBoundingClientRect()

    if (!slideRect || !blockRect || slideRect.width === 0 || slideRect.height === 0) {
      return fallback
    }

    return {
      x: ((blockRect.left - slideRect.left) / slideRect.width) * SLIDE_WIDTH,
      y: ((blockRect.top - slideRect.top) / slideRect.height) * SLIDE_HEIGHT,
      width: (blockRect.width / slideRect.width) * SLIDE_WIDTH,
      height: (blockRect.height / slideRect.height) * SLIDE_HEIGHT,
    }
  }, [slideRef])

  const calculateInteractionState = useCallback(
    (nextPoint: Point, currentInteraction: Interaction) => {
      const dx = nextPoint.x - currentInteraction.startPoint.x
      const dy = nextPoint.y - currentInteraction.startPoint.y

      if (currentInteraction.kind === 'move') {
        return calculateMoveInteractionState(
          dx,
          dy,
          currentInteraction,
          activeSlideBlocks,
          activeSlideIndex,
        )
      }

      const rect = resizeRect(
        currentInteraction.startRect,
        currentInteraction.handle,
        dx,
        dy,
      )

      return {
        guides: guidesForInteraction(rect, currentInteraction),
        rects: [{ pointer: currentInteraction.pointer, rect }],
      }
    },
    [activeSlideBlocks, activeSlideIndex],
  )

  useEffect(() => {
    if (!interaction) {
      return
    }

    const currentInteraction = interaction

    function handlePointerMove(event: PointerEvent) {
      const point = readSlidePoint(event)

      if (!point) {
        return
      }

      if (
        !hasMeaningfulPointerDelta(currentInteraction.startClientPoint, {
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        return
      }

      const { guides, rects } = calculateInteractionState(point, currentInteraction)
      setDraftLayout({
        rects,
      })
      setSnapGuides(guides)
    }

    function handlePointerUp(event: PointerEvent) {
      const point = readSlidePoint(event)

      if (
        !point ||
        !hasMeaningfulPointerDelta(currentInteraction.startClientPoint, {
          x: event.clientX,
          y: event.clientY,
        })
      ) {
        suppressStageClickRef.current = true
        clearLayoutInteraction()
        return
      }

      const { rects } = calculateInteractionState(point, currentInteraction)

      if (draftRectsEqual(rects, currentInteraction.startRects)) {
        suppressStageClickRef.current = true
        clearLayoutInteraction()
        return
      }

      suppressBlockClickRef.current = true
      window.setTimeout(() => {
        suppressBlockClickRef.current = false
      }, 0)
      const selection =
        currentInteraction.kind === 'move'
          ? selectionSnapForPointers(
              currentInteraction.pointers,
              currentInteraction.pointer,
            )
          : undefined

      suppressStageClickRef.current = true
      commitPatch(
        rects.flatMap(({ pointer, rect }) =>
          setArrangePatch(pointer, rect, {
            includeHeight: currentInteraction.kind === 'resize',
          }),
        ),
        currentInteraction.pointer,
        `${currentInteraction.kind} layout`,
        undefined,
        selection,
      )
      clearLayoutInteraction()
    }

    window.addEventListener('pointermove', handlePointerMove)
    window.addEventListener('pointerup', handlePointerUp)
    window.addEventListener('pointercancel', handlePointerUp)

    return () => {
      window.removeEventListener('pointermove', handlePointerMove)
      window.removeEventListener('pointerup', handlePointerUp)
      window.removeEventListener('pointercancel', handlePointerUp)
    }
  }, [
    calculateInteractionState,
    clearLayoutInteraction,
    commitPatch,
    interaction,
    readSlidePoint,
    suppressBlockClickRef,
    suppressStageClickRef,
  ])

  function startMoveInteraction(
    event: ReactPointerEvent<HTMLElement>,
    pointer: Pointer,
    block: SlideBlock,
  ) {
    if (mode !== 'layout') {
      return
    }

    const point = readSlidePoint(event)

    if (!point) {
      return
    }

    if (hasSelectionModifier(event)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    const activePointers = selectedPointerSet.has(pointer)
      ? selectedPointers
      : [pointer]
    const startRects = activePointers
      .map((activePointer) => {
        const location = blockLocationFromPointer(deckValue, activePointer)

        if (!location || location.slide.id !== activeSlideId) {
          return null
        }

        return {
          pointer: activePointer,
          rect: getCurrentRect(activePointer, location.block, draftLayout),
        }
      })
      .filter((draftRect): draftRect is DraftLayoutRect => draftRect !== null)

    if (!selectedPointerSet.has(pointer)) {
      selectBlock(pointer)
    }

    setInteraction({
      kind: 'move',
      pointer,
      pointers: startRects.map((draftRect) => draftRect.pointer),
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: getCurrentRect(pointer, block, draftLayout),
      startRects,
    })
  }

  function startResizeInteraction(
    event: ReactPointerEvent<HTMLButtonElement>,
    handle: ResizeHandle,
  ) {
    const currentSelectedRect =
      selectedPointer && selectedBlock
        ? readBlockVisualRect(
            selectedBlock.id,
            getCurrentRect(selectedPointer, selectedBlock, draftLayout),
          )
        : null

    if (
      selectedPointers.length !== 1 ||
      !selectedBlock ||
      !selectedPointer ||
      !currentSelectedRect ||
      mode !== 'layout'
    ) {
      return
    }

    const point = readSlidePoint(event)

    if (!point) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    setInteraction({
      kind: 'resize',
      pointer: selectedPointer,
      handle,
      startClientPoint: {
        x: event.clientX,
        y: event.clientY,
      },
      startPoint: point,
      startRect: currentSelectedRect,
      startRects: [{ pointer: selectedPointer, rect: currentSelectedRect }],
    })
  }

  return {
    clearLayoutInteraction,
    draftLayout,
    interaction,
    snapGuides,
    startMoveInteraction,
    startResizeInteraction,
  }
}
