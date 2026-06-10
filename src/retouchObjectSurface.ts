import {
  type ObjectSurfaceAdapter,
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
