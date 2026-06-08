import {
  appendSegment,
  buildPointer,
  parsePointer,
  type JSONPatchOperation,
  type Pointer,
} from 'zod-crud'
import {
  RESIZE_HANDLES as CANVAS_RESIZE_HANDLES,
  clamp as clampNumber,
  type ResizeHandle as CanvasResizeHandle,
} from 'canvas/core'
import { z } from 'zod'

export const SLIDE_WIDTH = 1280
export const SLIDE_HEIGHT = 720
export const GRID_SIZE = 8
export const MIN_BLOCK_SIZE = 72
export const EMPTY_TEXT_BOX_HEIGHT = 32

const ALIGN_SNAP_DISTANCE = GRID_SIZE

export const BlockRoleSchema = z.enum([
  'title',
  'subtitle',
  'body',
  'card',
  'metric',
  'chart',
  'note',
])
export const BlockTagSchema = z.enum(['h1', 'p', 'div'])

export const RectSchema = z.object({
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

export const RetouchTextPatchSchema = z.object({
  slideId: z.string(),
  blockId: z.string(),
  text: z.string(),
})

export const RetouchLayoutPatchSchema = z.object({
  slideId: z.string(),
  blockId: z.string(),
  rect: RectSchema,
})

export const RetouchPatchManifestSchema = z.object({
  version: z.literal(1),
  text: z.array(RetouchTextPatchSchema),
  layout: z.array(RetouchLayoutPatchSchema),
})

export type BlockRole = z.infer<typeof BlockRoleSchema>
export type BlockTag = z.infer<typeof BlockTagSchema>
export type Rect = z.infer<typeof RectSchema>
export type SlideBlock = z.infer<typeof SlideBlockSchema>
export type RetouchSlide = z.infer<typeof RetouchSlideSchema>
export type RetouchDeck = z.infer<typeof RetouchDeckSchema>
export type RetouchPatchManifest = z.infer<typeof RetouchPatchManifestSchema>

export type ResizeHandle = CanvasResizeHandle

export const RESIZE_HANDLES = CANVAS_RESIZE_HANDLES

export { SAMPLE_DECK, SAMPLE_SLIDES } from './sampleDeck'

export type BlockLocation = {
  pointer: Pointer
  slideIndex: number
  blockIndex: number
  slide: RetouchSlide
  block: SlideBlock
}

export function slidePointer(slideIndex: number): Pointer {
  return buildPointer(['slides', slideIndex])
}

export function slideBlocksPointer(slideIndex: number): Pointer {
  return appendSegment(slidePointer(slideIndex), 'blocks')
}

export function slideNamePointer(slideIndex: number): Pointer {
  return appendSegment(slidePointer(slideIndex), 'name')
}

export function slideAccentPointer(slideIndex: number): Pointer {
  return appendSegment(slidePointer(slideIndex), 'accent')
}

export function blockPointer(slideIndex: number, blockIndex: number): Pointer {
  return buildPointer(['slides', slideIndex, 'blocks', blockIndex])
}

export function blockTextPointer(pointer: Pointer): Pointer {
  return appendSegment(pointer, 'text')
}

export function findSlideIndex(deck: RetouchDeck, slideId: string) {
  return deck.slides.findIndex((slide) => slide.id === slideId)
}

export function findBlockLocation(
  deck: RetouchDeck,
  slideId: string,
  blockId: string,
): BlockLocation | null {
  const slideIndex = findSlideIndex(deck, slideId)

  if (slideIndex < 0) {
    return null
  }

  const slide = deck.slides[slideIndex]
  const blockIndex = slide.blocks.findIndex((block) => block.id === blockId)

  if (blockIndex < 0) {
    return null
  }

  return {
    pointer: blockPointer(slideIndex, blockIndex),
    slideIndex,
    blockIndex,
    slide,
    block: slide.blocks[blockIndex],
  }
}

export function blockLocationFromPointer(
  deck: RetouchDeck,
  pointer: Pointer | null | undefined,
): BlockLocation | null {
  if (!pointer) {
    return null
  }

  const segments = parsePointer(pointer)

  if (
    segments.length !== 4 ||
    segments[0] !== 'slides' ||
    segments[2] !== 'blocks'
  ) {
    return null
  }

  const slideIndex = Number(segments[1])
  const blockIndex = Number(segments[3])

  if (!Number.isInteger(slideIndex) || !Number.isInteger(blockIndex)) {
    return null
  }

  const slide = deck.slides[slideIndex]
  const block = slide?.blocks[blockIndex]

  if (!slide || !block) {
    return null
  }

  return {
    pointer,
    slideIndex,
    blockIndex,
    slide,
    block,
  }
}

export function blockLocationsFromPointers(
  deck: RetouchDeck,
  pointers: readonly Pointer[],
): BlockLocation[] {
  return pointers
    .map((pointer) => blockLocationFromPointer(deck, pointer))
    .filter((location): location is BlockLocation => location !== null)
    .sort((a, b) => a.slideIndex - b.slideIndex || a.blockIndex - b.blockIndex)
}

export function getRect(block: SlideBlock): Rect {
  return {
    x: block.x,
    y: block.y,
    width: block.width,
    height: block.height,
  }
}

export function setTextPatch(
  pointer: Pointer,
  text: string,
): JSONPatchOperation[] {
  return [{ op: 'replace', path: blockTextPointer(pointer), value: text }]
}

export function setLayoutPatch(
  pointer: Pointer,
  rect: Rect,
): JSONPatchOperation[] {
  return [
    { op: 'replace', path: appendSegment(pointer, 'x'), value: rect.x },
    { op: 'replace', path: appendSegment(pointer, 'y'), value: rect.y },
    { op: 'replace', path: appendSegment(pointer, 'width'), value: rect.width },
    { op: 'replace', path: appendSegment(pointer, 'height'), value: rect.height },
  ]
}

export function setArrangePatch(
  pointer: Pointer,
  rect: Rect,
  options: { includeHeight?: boolean } = {},
): JSONPatchOperation[] {
  const patch: JSONPatchOperation[] = [
    { op: 'replace', path: appendSegment(pointer, 'x'), value: rect.x },
    { op: 'replace', path: appendSegment(pointer, 'y'), value: rect.y },
    { op: 'replace', path: appendSegment(pointer, 'width'), value: rect.width },
  ]

  if (options.includeHeight) {
    patch.push({
      op: 'replace',
      path: appendSegment(pointer, 'height'),
      value: rect.height,
    })
  }

  return patch
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

export function clamp(value: number, min: number, max: number) {
  return clampNumber(value, min, max)
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
