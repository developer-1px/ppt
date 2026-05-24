import type { RetouchSlide } from './retouchModel'

const SLIDE_ACCENTS = [
  '#2563eb',
  '#0f766e',
  '#4f46e5',
  '#c2410c',
  '#7c3aed',
  '#0891b2',
  '#be123c',
  '#15803d',
]

export function createBlankSlide(slides: RetouchSlide[]): RetouchSlide {
  const ordinal = nextSlideOrdinal(slides)
  const id = uniqueSlideId(slides, `slide-${ordinal}`)

  return {
    id,
    name: uniqueSlideName(slides, `Slide ${ordinal}`),
    accent: SLIDE_ACCENTS[(ordinal - 1) % SLIDE_ACCENTS.length],
    blocks: [
      {
        id: `${id}-title`,
        role: 'title',
        tag: 'h1',
        text: 'Title',
        className: 'block-title compact centered',
        x: 300,
        y: 176,
        width: 680,
        height: 78,
      },
      {
        id: `${id}-subtitle`,
        role: 'subtitle',
        tag: 'p',
        text: 'Subtitle',
        className: 'block-subtitle centered',
        x: 300,
        y: 300,
        width: 680,
        height: 72,
      },
      {
        id: `${id}-note`,
        role: 'note',
        tag: 'p',
        text: 'Key point',
        className: 'block-note slim centered',
        x: 260,
        y: 552,
        width: 760,
        height: 76,
      },
    ],
  }
}

export function duplicateSlide(
  slide: RetouchSlide,
  slides: RetouchSlide[],
): RetouchSlide {
  const ordinal = nextSlideOrdinal(slides)
  const id = uniqueSlideId(slides, `slide-${ordinal}`)

  return {
    ...slide,
    id,
    name: uniqueSlideName(slides, `${slide.name} copy`),
    blocks: slide.blocks.map((block, index) => ({
      ...block,
      id: `${id}-block-${index + 1}`,
    })),
  }
}

function nextSlideOrdinal(slides: RetouchSlide[]) {
  const maxOrdinal = slides.reduce((max, slide) => {
    const match = /^slide-(\d+)$/.exec(slide.id)

    return match ? Math.max(max, Number(match[1])) : max
  }, 0)

  return Math.max(maxOrdinal, slides.length) + 1
}

function uniqueSlideId(slides: RetouchSlide[], baseId: string) {
  const existingIds = new Set(slides.map((slide) => slide.id))

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

function uniqueSlideName(slides: RetouchSlide[], baseName: string) {
  const existingNames = new Set(slides.map((slide) => slide.name))

  if (!existingNames.has(baseName)) {
    return baseName
  }

  for (let suffix = 2; ; suffix += 1) {
    const candidate = `${baseName} ${suffix}`

    if (!existingNames.has(candidate)) {
      return candidate
    }
  }
}
