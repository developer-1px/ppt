import {
  rectFromPoints,
  selectSurfaceObjectsInMarquee,
} from '@interactive-os/object-surface'
import {
  useCallback,
  useEffect,
  useState,
  type MutableRefObject,
  type PointerEvent as ReactPointerEvent,
  type RefObject,
} from 'react'
import type { Pointer } from 'zod-crud'
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  type Rect,
  type RetouchSlide,
} from './retouchModel'
import { clamp, hasSelectionModifier, type Point } from './layoutInteraction'
import {
  objectSurfaceSelectionFromPointers,
  pointersFromObjectSurfaceSelection,
  retouchSurfaceAdapter,
  retouchSurfaceItems,
} from './retouchObjectSurface'

const MARQUEE_THRESHOLD = 6

type SelectionApi = {
  empty?: () => void
  selectRanges?: (pointers: Pointer[]) => void
}

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
  mode: 'text' | 'layout'
  selectedPointers: Pointer[]
  selection: SelectionApi | null | undefined
  slideRef: RefObject<HTMLDivElement | null>
  suppressStageClickRef: MutableRefObject<boolean>
}) {
  const [marquee, setMarquee] = useState<MarqueeState | null>(null)

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

  const startMarqueeSelection = useCallback(
    (event: ReactPointerEvent<HTMLDivElement>) => {
      if (mode !== 'layout' || event.button !== 0) {
        return
      }

      const target = event.target instanceof HTMLElement ? event.target : null

      if (target?.closest('[data-block], .selection-overlay, .resize-handle')) {
        return
      }

      const point = readSlidePoint(event)

      if (!point) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      setMarquee({
        additive: hasSelectionModifier(event),
        rect: null,
        startClientPoint: {
          x: event.clientX,
          y: event.clientY,
        },
        startPoint: point,
      })
    },
    [mode, readSlidePoint],
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
      const point = readSlidePoint(event)

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
      const point = readSlidePoint(event)

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
    readSlidePoint,
    selectedPointers,
    selection,
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
