import type { SlideBlock } from './retouchModel'
import { PPT_RETOUCH_SURFACE } from './retouchSurfaceContract'

const HTML_SLIDE_CONTRACT = 'interactive-os.html-slide-contract.v1' as const

export const HTML_SLIDE_ATTRIBUTES = {
  block: 'data-block',
  blockIndex: 'data-block-index',
  contract: 'data-html-slide-contract',
  role: 'data-role',
  slide: 'data-slide',
  surfaceContract: 'data-retouch-surface-contract',
} as const

export const HTML_SLIDE_CLASSES = {
  blockText: 'slide-block-text',
} as const

export type HtmlSlideRootAttributes = {
  [HTML_SLIDE_ATTRIBUTES.contract]: typeof HTML_SLIDE_CONTRACT
  [HTML_SLIDE_ATTRIBUTES.slide]: string
  [HTML_SLIDE_ATTRIBUTES.surfaceContract]: typeof PPT_RETOUCH_SURFACE.contract
}

export type HtmlSlideBlockAttributes = {
  [HTML_SLIDE_ATTRIBUTES.block]: string
  [HTML_SLIDE_ATTRIBUTES.blockIndex]: string
  [HTML_SLIDE_ATTRIBUTES.role]: SlideBlock['role']
}

export function htmlSlideRootAttributes(slideId: string): HtmlSlideRootAttributes {
  return {
    [HTML_SLIDE_ATTRIBUTES.contract]: HTML_SLIDE_CONTRACT,
    [HTML_SLIDE_ATTRIBUTES.slide]: slideId,
    [HTML_SLIDE_ATTRIBUTES.surfaceContract]: PPT_RETOUCH_SURFACE.contract,
  }
}

export function htmlSlideBlockAttributes(
  block: SlideBlock,
  blockIndex: number,
): HtmlSlideBlockAttributes {
  return {
    [HTML_SLIDE_ATTRIBUTES.block]: block.id,
    [HTML_SLIDE_ATTRIBUTES.blockIndex]: String(blockIndex),
    [HTML_SLIDE_ATTRIBUTES.role]: block.role,
  }
}
