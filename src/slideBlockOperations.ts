import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  clamp,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import {
  createRetouchBlockCopyId,
  createRetouchTextBlockId,
  nextRetouchTextBlockOrdinal,
} from './retouchIdResolver'

const DUPLICATE_OFFSET = 32

export function createTextBlock(slide: RetouchSlide): SlideBlock {
  const ordinal = nextRetouchTextBlockOrdinal(slide)
  const id = createRetouchTextBlockId(slide)
  const offset = ((ordinal - 1) % 5) * 24

  return {
    id,
    role: 'body',
    tag: 'p',
    text: '',
    className: 'block-subtitle',
    x: 360 + offset,
    y: 312 + offset,
    width: 560,
    height: 56,
  }
}

export function duplicateBlock(
  block: SlideBlock,
  slide: RetouchSlide,
): SlideBlock {
  return {
    ...block,
    id: createRetouchBlockCopyId(slide, block),
    x: clamp(block.x + DUPLICATE_OFFSET, 0, SLIDE_WIDTH - block.width),
    y: clamp(block.y + DUPLICATE_OFFSET, 0, SLIDE_HEIGHT - block.height),
  }
}
