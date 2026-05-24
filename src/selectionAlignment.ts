import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  clamp,
  type Rect,
} from './retouchModel'

export type AlignSelectionAction =
  | 'left'
  | 'center-x'
  | 'right'
  | 'top'
  | 'middle-y'
  | 'bottom'

export function alignRectToBounds(
  rect: Rect,
  bounds: Rect,
  action: AlignSelectionAction,
): Rect {
  let x = rect.x
  let y = rect.y

  if (action === 'left') {
    x = bounds.x
  }

  if (action === 'center-x') {
    x = bounds.x + bounds.width / 2 - rect.width / 2
  }

  if (action === 'right') {
    x = bounds.x + bounds.width - rect.width
  }

  if (action === 'top') {
    y = bounds.y
  }

  if (action === 'middle-y') {
    y = bounds.y + bounds.height / 2 - rect.height / 2
  }

  if (action === 'bottom') {
    y = bounds.y + bounds.height - rect.height
  }

  return {
    ...rect,
    x: clamp(x, 0, SLIDE_WIDTH - rect.width),
    y: clamp(y, 0, SLIDE_HEIGHT - rect.height),
  }
}

export function alignmentBounds(rects: Rect[]) {
  if (rects.length === 0) {
    return null
  }

  if (rects.length === 1) {
    return {
      x: 0,
      y: 0,
      width: SLIDE_WIDTH,
      height: SLIDE_HEIGHT,
    }
  }

  return rectBounds(rects)
}

export function rectBounds(rects: Rect[]) {
  if (rects.length === 0) {
    return null
  }

  const left = Math.min(...rects.map((rect) => rect.x))
  const top = Math.min(...rects.map((rect) => rect.y))
  const right = Math.max(...rects.map((rect) => rect.x + rect.width))
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.height))

  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  }
}
