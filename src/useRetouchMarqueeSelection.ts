import {
  rectFromPoints,
  selectSurfaceObjectsInMarquee,
} from '@interactive-os/object-surface'
import { isAdditivePointerInput } from 'canvas/foundation'
import type { Point } from 'canvas/core'
import {
  useCallback,
  useEffect,
  useState,
  type MutableRefObject,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import type { Pointer, SelectionState } from 'zod-crud'
import {
  type Rect,
  type RetouchSlide,
} from './retouchModel'
import {
  objectSurfaceSelectionFromPointers,
  pointersFromObjectSurfaceSelection,
  retouchSurfaceAdapter,
  retouchSurfaceItems,
} from './retouchObjectSurface'
import { readSlidePoint } from './retouchSlideDom'
import type { RetouchMode } from './retouchViewState'

const MARQUEE_THRESHOLD = 6

type MarqueeState = {
  additive: boolean
  rect: Rect | null
  startClientPoint: Point
  startPoint: Point
}

export function useRetouchMarqueeSelection({
  activeSlide,
  activeSlideIndex,
  mode,
  selectedPointers,
  selection,
  slideRef,
  suppressStageClickRef,
}: {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  mode: RetouchMode
  selectedPointers: Pointer[]
  selection: SelectionState | null | undefined
  slideRef: RefObject<HTMLDivElement | null>
  suppressStageClickRef: MutableRefObject<boolean>
}) {
  const [marquee, setMarquee] = useState<MarqueeState | null>(null)

  const startMarqueeSelection = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (mode !== 'layout' || event.button !== 0) {
        return
      }

      const target = event.target instanceof HTMLElement ? event.target : null

      if (target?.closest('[data-block], .selection-overlay, .resize-handle')) {
        return
      }

      const point = readSlidePoint(slideRef.current, event)

      if (!point) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      setMarquee({
        additive: isAdditivePointerInput(event),
        rect: null,
        startClientPoint: {
          x: event.clientX,
          y: event.clientY,
        },
        startPoint: point,
      })
    },
    [mode, slideRef],
  )

  const clearMarqueeSelection = useCallback(() => {
    setMarquee(null)
  }, [])

  useEffect(() => {
    if (!marquee) {
      return
    }

    const currentMarquee = marquee

    function handlePointerMove(event: PointerEvent) {
      const point = readSlidePoint(slideRef.current, event)

      if (!point || !hasMeaningfulClientDelta(currentMarquee.startClientPoint, event)) {
        return
      }

      setMarquee((current) =>
        current
          ? {
              ...current,
              rect: rectFromPoints(current.startPoint, point),
            }
          : current,
      )
    }

    function handlePointerUp(event: PointerEvent) {
      const point = readSlidePoint(slideRef.current, event)

      if (!point || !hasMeaningfulClientDelta(currentMarquee.startClientPoint, event)) {
        setMarquee(null)
        return
      }

      const rect = rectFromPoints(currentMarquee.startPoint, point)
      const nextSurfaceSelection = selectSurfaceObjectsInMarquee({
        adapter: retouchSurfaceAdapter,
        items: retouchSurfaceItems(activeSlide, activeSlideIndex),
        mode: currentMarquee.additive ? 'add' : 'replace',
        rect,
        selection: objectSurfaceSelectionFromPointers(selectedPointers),
      })
      const nextPointers = pointersFromObjectSurfaceSelection(nextSurfaceSelection)

      if (nextPointers.length > 0) {
        selection?.selectRanges?.(nextPointers)
      } else if (!currentMarquee.additive) {
        selection?.empty?.()
      }

      suppressStageClickRef.current = true
      setMarquee(null)
    }

    window.addEventListener('pointermove', handlePointerMove)
    window.addEventListener('pointerup', handlePointerUp)
    window.addEventListener('pointercancel', clearMarqueeSelection)

    return () => {
      window.removeEventListener('pointermove', handlePointerMove)
      window.removeEventListener('pointerup', handlePointerUp)
      window.removeEventListener('pointercancel', clearMarqueeSelection)
    }
  }, [
    activeSlide,
    activeSlideIndex,
    clearMarqueeSelection,
    marquee,
    selectedPointers,
    selection,
    slideRef,
    suppressStageClickRef,
  ])

  return {
    clearMarqueeSelection,
    marqueeRect: marquee?.rect ?? null,
    startMarqueeSelection,
  }
}

function hasMeaningfulClientDelta(start: Point, next: Pick<PointerEvent, 'clientX' | 'clientY'>) {
  return (
    Math.abs(next.clientX - start.x) >= MARQUEE_THRESHOLD ||
    Math.abs(next.clientY - start.y) >= MARQUEE_THRESHOLD
  )
}
