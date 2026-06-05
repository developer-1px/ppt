import {
  createObjectSurfaceSelection,
  getSurfaceSelectionBounds,
  type ObjectSurfaceAdapter,
  type ObjectSurfaceSelection,
} from '@interactive-os/object-surface'
import type { Pointer } from 'zod-crud'
import {
  blockPointer,
  getRect,
  type Rect,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'

export type RetouchSurfaceItem = {
  block: SlideBlock
  pointer: Pointer
  rect: Rect
}

export const retouchSurfaceAdapter: ObjectSurfaceAdapter<RetouchSurfaceItem> = {
  getBounds: (item) => item.rect,
  getId: (item) => item.pointer,
}

export function retouchSurfaceItems(slide: RetouchSlide, slideIndex: number) {
  return slide.blocks.map((block, blockIndex) => ({
    block,
    pointer: blockPointer(slideIndex, blockIndex),
    rect: getRect(block),
  }))
}

export function objectSurfaceSelectionFromPointers(
  pointers: readonly Pointer[],
): ObjectSurfaceSelection {
  return createObjectSurfaceSelection(pointers, {
    anchorId: pointers[0] ?? null,
    primaryId: pointers.at(-1) ?? null,
  })
}

export function pointersFromObjectSurfaceSelection(
  selection: ObjectSurfaceSelection,
): Pointer[] {
  return [...selection.ids]
}

export function retouchSurfaceSelectionBounds(
  items: readonly RetouchSurfaceItem[],
  pointers: readonly Pointer[],
) {
  return getSurfaceSelectionBounds(
    items,
    retouchSurfaceAdapter,
    pointers,
  )
}
