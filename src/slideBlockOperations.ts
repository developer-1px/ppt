import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  clamp,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'

const DUPLICATE_OFFSET = 32

export function createTextBlock(slide: RetouchSlide): SlideBlock {
  const ordinal = nextTextBlockOrdinal(slide)
  const id = uniqueBlockId(slide, `${slide.id}-text-${ordinal}`)
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
    id: uniqueBlockId(slide, `${block.id}-copy`),
    x: clamp(block.x + DUPLICATE_OFFSET, 0, SLIDE_WIDTH - block.width),
    y: clamp(block.y + DUPLICATE_OFFSET, 0, SLIDE_HEIGHT - block.height),
  }
}

function nextTextBlockOrdinal(slide: RetouchSlide) {
  const maxOrdinal = slide.blocks.reduce((max, block) => {
    const match = new RegExp(`^${escapeRegExp(slide.id)}-text-(\\d+)$`).exec(
      block.id,
    )

    return match ? Math.max(max, Number(match[1])) : max
  }, 0)

  return maxOrdinal + 1
}

function uniqueBlockId(slide: RetouchSlide, baseId: string) {
  const existingIds = new Set(slide.blocks.map((block) => block.id))

  if (!existingIds.has(baseId)) {
    return baseId
  }

  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${baseId}-${suffix}`

    if (!existingIds.has(candidate)) {
      return candidate
    }
  }
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
