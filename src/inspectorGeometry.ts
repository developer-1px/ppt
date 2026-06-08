import { clamp } from 'canvas/core'
import {
  MIN_BLOCK_SIZE,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  snap,
  type Rect,
} from './retouchModel'

export type RectField = keyof Rect

export function normalizeInspectorRect(
  rect: Rect,
  currentRect: Rect,
  changedField?: RectField,
): Rect {
  const nextRect = { ...currentRect }

  if (changedField === 'x') {
    nextRect.x = clamp(
      snap(finiteNumber(rect.x, currentRect.x)),
      0,
      SLIDE_WIDTH - currentRect.width,
    )
  }

  if (changedField === 'y') {
    nextRect.y = clamp(
      snap(finiteNumber(rect.y, currentRect.y)),
      0,
      SLIDE_HEIGHT - currentRect.height,
    )
  }

  if (changedField === 'width') {
    nextRect.width = clamp(
      snap(finiteNumber(rect.width, currentRect.width)),
      MIN_BLOCK_SIZE,
      SLIDE_WIDTH - currentRect.x,
    )
  }

  if (changedField === 'height') {
    nextRect.height = clamp(
      snap(finiteNumber(rect.height, currentRect.height)),
      MIN_BLOCK_SIZE,
      SLIDE_HEIGHT - currentRect.y,
    )
  }

  return nextRect
}

function finiteNumber(value: number, fallback: number) {
  return Number.isFinite(value) ? value : fallback
}
