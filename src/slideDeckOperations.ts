import type { JSONPatchOperation } from 'zod-crud'
import {
  slidePointer,
  type RetouchSlide,
} from './retouchModel'
import {
  createRetouchSlideId,
  createRetouchSlideName,
  nextRetouchSlideOrdinal,
} from './retouchIdResolver'

export const SLIDE_ACCENTS = [
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
  const ordinal = nextRetouchSlideOrdinal(slides)
  const id = createRetouchSlideId(slides, ordinal)

  return {
    id,
    name: createRetouchSlideName(slides, `Slide ${ordinal}`),
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
  const ordinal = nextRetouchSlideOrdinal(slides)
  const id = createRetouchSlideId(slides, ordinal)

  return {
    ...slide,
    id,
    name: createRetouchSlideName(slides, `${slide.name} copy`),
    blocks: slide.blocks.map((block, index) => ({
      ...block,
      id: `${id}-block-${index + 1}`,
    })),
  }
}

export function addSlidePatch(
  slide: RetouchSlide,
  insertIndex: number,
): JSONPatchOperation[] {
  return [{ op: 'add', path: slidePointer(insertIndex), value: slide }]
}
