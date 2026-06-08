export type CanvasZoom = 'fit' | number

const NORMAL_CANVAS_ZOOM = 1
const CANVAS_ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2]
const MIN_CANVAS_ZOOM = CANVAS_ZOOM_STEPS[0] ?? NORMAL_CANVAS_ZOOM
const MAX_CANVAS_ZOOM =
  CANVAS_ZOOM_STEPS[CANVAS_ZOOM_STEPS.length - 1] ?? NORMAL_CANVAS_ZOOM

export function nextCanvasZoom(
  canvasZoom: CanvasZoom,
  direction: -1 | 1,
): CanvasZoom {
  const currentZoom = canvasZoom === 'fit' ? NORMAL_CANVAS_ZOOM : canvasZoom
  const currentIndex = CANVAS_ZOOM_STEPS.findIndex((step) => step >= currentZoom)
  const fallbackIndex = CANVAS_ZOOM_STEPS.indexOf(NORMAL_CANVAS_ZOOM)
  const nextIndex =
    direction > 0
      ? currentZoom >= MAX_CANVAS_ZOOM
        ? CANVAS_ZOOM_STEPS.length - 1
        : Math.max(0, currentIndex) + 1
      : currentZoom <= MIN_CANVAS_ZOOM
        ? 0
        : currentIndex > 0
          ? currentIndex - 1
          : fallbackIndex

  return CANVAS_ZOOM_STEPS[nextIndex] ?? NORMAL_CANVAS_ZOOM
}

export function canIncreaseCanvasZoom(canvasZoom: CanvasZoom) {
  return canvasZoom === 'fit' || canvasZoom < MAX_CANVAS_ZOOM
}

export function canDecreaseCanvasZoom(canvasZoom: CanvasZoom) {
  return canvasZoom === 'fit' || canvasZoom > MIN_CANVAS_ZOOM
}

export function canvasZoomLabel(canvasZoom: CanvasZoom) {
  return canvasZoom === 'fit' ? 'Fit' : `${Math.round(canvasZoom * 100)}%`
}
