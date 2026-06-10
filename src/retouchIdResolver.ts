import { createIdResolver } from '@zod-crud/id-resolver'
import type { ResolveIdResult } from '@zod-crud/id-resolver'
import { parsePointer, type JSONDocument, type Pointer } from 'zod-crud'
import {
  RetouchSlideSchema,
  SlideBlockSchema,
  type RetouchDeck,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'

const RETOUCH_ID_SCOPES = {
  block: 'block',
  slide: 'slide',
} as const

export function createRetouchIdResolver(doc: JSONDocument<RetouchDeck>) {
  const resolver = createIdResolver(doc, {
    scopes: [
      {
        scope: RETOUCH_ID_SCOPES.slide,
        query: '$.slides[*]',
        readId: (value) => RetouchSlideSchema.safeParse(value).data?.id,
      },
      {
        scope: RETOUCH_ID_SCOPES.block,
        query: '$.slides[*].blocks[*]',
        readId: (value) => SlideBlockSchema.safeParse(value).data?.id,
      },
    ],
  })

  return {
    current: resolver.current,

    resolveBlockPointer(blockId: string): ResolveIdResult {
      return resolver.resolve(RETOUCH_ID_SCOPES.block, blockId)
    },

    resolveSlideIndex(slideId: string): number | null {
      const resolved = resolver.resolve(RETOUCH_ID_SCOPES.slide, slideId)

      return resolved.ok ? slideIndexFromPointer(resolved.pointer) : null
    },

    resolveSlidePointer(slideId: string): ResolveIdResult {
      return resolver.resolve(RETOUCH_ID_SCOPES.slide, slideId)
    },
  }
}

export function nextRetouchSlideOrdinal(slides: readonly RetouchSlide[]) {
  const maxOrdinal = slides.reduce((max, slide) => {
    const match = /^slide-(\d+)$/.exec(slide.id)

    return match ? Math.max(max, Number(match[1])) : max
  }, 0)

  return Math.max(maxOrdinal, slides.length) + 1
}

export function createRetouchSlideId(
  slides: readonly RetouchSlide[],
  ordinal = nextRetouchSlideOrdinal(slides),
) {
  return uniqueId(
    slides.map((slide) => slide.id),
    `slide-${ordinal}`,
  )
}

export function createRetouchSlideName(
  slides: readonly RetouchSlide[],
  baseName: string,
) {
  return uniqueId(
    slides.map((slide) => slide.name),
    baseName,
    ' ',
  )
}

export function nextRetouchTextBlockOrdinal(slide: RetouchSlide) {
  const maxOrdinal = slide.blocks.reduce((max, block) => {
    const match = new RegExp(`^${escapeRegExp(slide.id)}-text-(\\d+)$`).exec(
      block.id,
    )

    return match ? Math.max(max, Number(match[1])) : max
  }, 0)

  return maxOrdinal + 1
}

export function createRetouchBlockCopyId(
  slide: RetouchSlide,
  block: SlideBlock,
) {
  return uniqueId(
    slide.blocks.map((block) => block.id),
    `${block.id}-copy`,
  )
}

function slideIndexFromPointer(pointer: Pointer) {
  const segments = parsePointer(pointer)

  if (segments.length !== 2 || segments[0] !== 'slides') {
    return null
  }

  const slideIndex = Number(segments[1])

  return Number.isInteger(slideIndex) ? slideIndex : null
}

function uniqueId(
  existingValues: readonly string[],
  baseValue: string,
  suffixSeparator = '-',
) {
  const existing = new Set(existingValues)

  if (!existing.has(baseValue)) {
    return baseValue
  }

  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${baseValue}${suffixSeparator}${suffix}`

    if (!existing.has(candidate)) {
      return candidate
    }
  }
}

function escapeRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}
