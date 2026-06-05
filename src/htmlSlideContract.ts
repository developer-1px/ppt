import type { SlideBlock } from './retouchModel'
import { RETOUCH_SURFACE_CONTRACT } from './retouchSurfaceContract'

export const HTML_SLIDE_CONTRACT = 'interactive-os.html-slide-contract.v1' as const

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

export const HTML_SLIDE_REQUIRED_ROOT_ATTRIBUTES = [
  HTML_SLIDE_ATTRIBUTES.contract,
  HTML_SLIDE_ATTRIBUTES.slide,
  HTML_SLIDE_ATTRIBUTES.surfaceContract,
] as const

export const HTML_SLIDE_REQUIRED_BLOCK_ATTRIBUTES = [
  HTML_SLIDE_ATTRIBUTES.block,
  HTML_SLIDE_ATTRIBUTES.blockIndex,
  HTML_SLIDE_ATTRIBUTES.role,
] as const

export type HtmlSlideRootAttributes = {
  [HTML_SLIDE_ATTRIBUTES.contract]: typeof HTML_SLIDE_CONTRACT
  [HTML_SLIDE_ATTRIBUTES.slide]: string
  [HTML_SLIDE_ATTRIBUTES.surfaceContract]: typeof RETOUCH_SURFACE_CONTRACT
}

export type HtmlSlideBlockAttributes = {
  [HTML_SLIDE_ATTRIBUTES.block]: string
  [HTML_SLIDE_ATTRIBUTES.blockIndex]: string
  [HTML_SLIDE_ATTRIBUTES.role]: SlideBlock['role']
}

export type HtmlSlideBlockSnapshot = {
  blockId: string
  blockIndex: number
  boundsStyle: {
    height: string
    left: string
    minHeight: string
    top: string
    width: string
  }
  role: string
}

export type HtmlSlideSnapshot = {
  blocks: HtmlSlideBlockSnapshot[]
  contract: string
  slideId: string
  surfaceContract: string
}

export function htmlSlideRootAttributes(slideId: string): HtmlSlideRootAttributes {
  return {
    [HTML_SLIDE_ATTRIBUTES.contract]: HTML_SLIDE_CONTRACT,
    [HTML_SLIDE_ATTRIBUTES.slide]: slideId,
    [HTML_SLIDE_ATTRIBUTES.surfaceContract]: RETOUCH_SURFACE_CONTRACT,
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

export function readHtmlSlideSnapshot(slideElement: HTMLElement): HtmlSlideSnapshot {
  const blocks = Array.from(
    slideElement.querySelectorAll<HTMLElement>(
      `[${HTML_SLIDE_ATTRIBUTES.block}]`,
    ),
  ).map((element, fallbackIndex) => {
    const blockIndex = Number(
      element.getAttribute(HTML_SLIDE_ATTRIBUTES.blockIndex) ?? fallbackIndex,
    )

    return {
      blockId: element.getAttribute(HTML_SLIDE_ATTRIBUTES.block) ?? '',
      blockIndex: Number.isFinite(blockIndex) ? blockIndex : fallbackIndex,
      boundsStyle: {
        height: element.style.height,
        left: element.style.left,
        minHeight: element.style.minHeight,
        top: element.style.top,
        width: element.style.width,
      },
      role: element.getAttribute(HTML_SLIDE_ATTRIBUTES.role) ?? '',
    }
  })

  return {
    blocks,
    contract: slideElement.getAttribute(HTML_SLIDE_ATTRIBUTES.contract) ?? '',
    slideId: slideElement.getAttribute(HTML_SLIDE_ATTRIBUTES.slide) ?? '',
    surfaceContract:
      slideElement.getAttribute(HTML_SLIDE_ATTRIBUTES.surfaceContract) ?? '',
  }
}

export function validateHtmlSlideSnapshot(snapshot: HtmlSlideSnapshot): string[] {
  const errors: string[] = []

  if (snapshot.contract !== HTML_SLIDE_CONTRACT) {
    errors.push('html slide contract mismatch')
  }

  if (snapshot.surfaceContract !== RETOUCH_SURFACE_CONTRACT) {
    errors.push('retouch surface contract mismatch')
  }

  if (!snapshot.slideId) {
    errors.push('missing slide id')
  }

  snapshot.blocks.forEach((block, index) => {
    if (!block.blockId) {
      errors.push(`missing block id at index ${index}`)
    }

    if (!block.role) {
      errors.push(`missing block role at index ${index}`)
    }

    if (block.blockIndex !== index) {
      errors.push(`block index mismatch at index ${index}`)
    }

    if (!block.boundsStyle.left || !block.boundsStyle.top || !block.boundsStyle.width) {
      errors.push(`missing block bounds at index ${index}`)
    }
  })

  return errors
}
