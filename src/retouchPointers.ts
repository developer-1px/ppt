import {
  appendSegment,
  buildPointer,
  parsePointer,
  type JSONPatchOperation,
  type Pointer,
} from 'zod-crud'
import type {
  Rect,
  RetouchDeck,
  RetouchSlide,
  SlideBlock,
} from './retouchModel'

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

export function findBlockLocation(
  deck: RetouchDeck,
  slideId: string,
  blockId: string,
): BlockLocation | null {
  const slideIndex = deck.slides.findIndex((slide) => slide.id === slideId)

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
  if (pointer === null || pointer === undefined) {
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

  if (slide === undefined || block === undefined) {
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
