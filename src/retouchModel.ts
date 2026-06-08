import { clamp, type ResizeHandle } from 'canvas/core'
import { z } from 'zod'
export {
  blockLocationFromPointer,
  blockLocationsFromPointers,
  blockPointer,
  blockTextPointer,
  findBlockLocation,
  setArrangePatch,
  setLayoutPatch,
  setTextPatch,
  slideAccentPointer,
  slideBlocksPointer,
  slideNamePointer,
  slidePointer,
  type BlockLocation,
} from './retouchPointers'

export const SLIDE_WIDTH = 1280
export const SLIDE_HEIGHT = 720
export const GRID_SIZE = 8
export const MIN_BLOCK_SIZE = 72
export const EMPTY_TEXT_BOX_HEIGHT = 32

const ALIGN_SNAP_DISTANCE = GRID_SIZE

const BlockRoleSchema = z.enum([
  'title',
  'subtitle',
  'body',
  'card',
  'metric',
  'chart',
  'note',
])
const BlockTagSchema = z.enum(['h1', 'p', 'div'])

const RectSchema = z.object({
  x: z.number(),
  y: z.number(),
  width: z.number(),
  height: z.number(),
})

export const SlideBlockSchema = RectSchema.extend({
  id: z.string(),
  role: BlockRoleSchema,
  tag: BlockTagSchema,
  text: z.string(),
  className: z.string(),
})

export const RetouchSlideSchema = z.object({
  id: z.string(),
  name: z.string(),
  accent: z.string(),
  blocks: z.array(SlideBlockSchema),
})

export const RetouchDeckSchema = z.object({
  slides: z.array(RetouchSlideSchema),
})

export type Rect = z.infer<typeof RectSchema>
export type SlideBlock = z.infer<typeof SlideBlockSchema>
export type RetouchSlide = z.infer<typeof RetouchSlideSchema>
export type RetouchDeck = z.infer<typeof RetouchDeckSchema>
export type RetouchPatchManifest = {
  version: 1
  text: {
    slideId: string
    blockId: string
    text: string
  }[]
  layout: {
    slideId: string
    blockId: string
    rect: Rect
  }[]
}

export function getRect(block: SlideBlock): Rect {
  return {
    x: block.x,
    y: block.y,
    width: block.width,
    height: block.height,
  }
}

export function rectToStyle(rect: Rect) {
  return {
    left: `${(rect.x / SLIDE_WIDTH) * 100}%`,
    top: `${(rect.y / SLIDE_HEIGHT) * 100}%`,
    width: `${(rect.width / SLIDE_WIDTH) * 100}%`,
    height: `${(rect.height / SLIDE_HEIGHT) * 100}%`,
  }
}

export function rectToAutoHeightStyle(rect: Rect, minimumHeight: number) {
  return {
    left: `${(rect.x / SLIDE_WIDTH) * 100}%`,
    top: `${(rect.y / SLIDE_HEIGHT) * 100}%`,
    width: `${(rect.width / SLIDE_WIDTH) * 100}%`,
    height: 'auto',
    minHeight:
      minimumHeight > 0
        ? `${(minimumHeight / SLIDE_HEIGHT) * 100}%`
        : undefined,
  }
}

export function snap(value: number) {
  return Math.round(value / GRID_SIZE) * GRID_SIZE
}

export function moveRect(rect: Rect, dx: number, dy: number): Rect {
  const x = snapAlignedCenter(
    dx === 0 ? rect.x : snap(rect.x + dx),
    rect.width,
    SLIDE_WIDTH,
    dx !== 0,
  )
  const y = snapAlignedCenter(
    dy === 0 ? rect.y : snap(rect.y + dy),
    rect.height,
    SLIDE_HEIGHT,
    dy !== 0,
  )

  return {
    ...rect,
    x: clamp(x, 0, SLIDE_WIDTH - rect.width),
    y: clamp(y, 0, SLIDE_HEIGHT - rect.height),
  }
}

export function resizeRect(
  rect: Rect,
  handle: ResizeHandle,
  dx: number,
  dy: number,
): Rect {
  let left = rect.x
  let right = rect.x + rect.width
  let top = rect.y
  let bottom = rect.y + rect.height

  if (handle.includes('w')) {
    left = clamp(snap(left + dx), 0, right - MIN_BLOCK_SIZE)
  }

  if (handle.includes('e')) {
    right = clamp(snap(right + dx), left + MIN_BLOCK_SIZE, SLIDE_WIDTH)
  }

  if (handle.includes('n')) {
    top = clamp(snap(top + dy), 0, bottom - MIN_BLOCK_SIZE)
  }

  if (handle.includes('s')) {
    bottom = clamp(snap(bottom + dy), top + MIN_BLOCK_SIZE, SLIDE_HEIGHT)
  }

  return {
    ...rect,
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  }
}

export function rectEquals(a: Rect, b: Rect) {
  return (
    a.x === b.x &&
    a.y === b.y &&
    a.width === b.width &&
    a.height === b.height
  )
}

function snapAlignedCenter(
  start: number,
  size: number,
  containerSize: number,
  enabled: boolean,
) {
  if (!enabled) {
    return start
  }

  const center = start + size / 2
  const target = containerSize / 2

  return Math.abs(center - target) <= ALIGN_SNAP_DISTANCE
    ? target - size / 2
    : start
}
