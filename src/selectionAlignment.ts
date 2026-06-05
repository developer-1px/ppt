import { unionRects } from '@interactive-os/object-surface'
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

export type DistributeSelectionAction = 'horizontal' | 'vertical'

export type DistributionTarget<T> = {
  item: T
  rect: Rect
}

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

export function distributeRects<T>(
  targets: DistributionTarget<T>[],
  action: DistributeSelectionAction,
) {
  if (targets.length < 3) {
    return targets
  }

  const bounds = rectBounds(targets.map((target) => target.rect))

  if (!bounds) {
    return targets
  }

  const sortedTargets = [...targets].sort((a, b) =>
    action === 'horizontal'
      ? a.rect.x - b.rect.x || a.rect.y - b.rect.y
      : a.rect.y - b.rect.y || a.rect.x - b.rect.x,
  )
  const totalSize = sortedTargets.reduce(
    (sum, target) =>
      sum + (action === 'horizontal' ? target.rect.width : target.rect.height),
    0,
  )
  const gap =
    ((action === 'horizontal' ? bounds.width : bounds.height) - totalSize) /
    (sortedTargets.length - 1)
  let cursor = action === 'horizontal' ? bounds.x : bounds.y

  return sortedTargets.map((target) => {
    const rect =
      action === 'horizontal'
        ? {
            ...target.rect,
            x: cursor,
          }
        : {
            ...target.rect,
            y: cursor,
          }

    cursor += (action === 'horizontal' ? target.rect.width : target.rect.height) + gap

    return {
      ...target,
      rect,
    }
  })
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
  return unionRects(rects)
}
