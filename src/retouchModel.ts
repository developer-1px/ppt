import {
  appendSegment,
  buildPointer,
  parsePointer,
  type JSONPatchOperation,
  type Pointer,
} from 'zod-crud'
import { z } from 'zod'

export const SLIDE_WIDTH = 1280
export const SLIDE_HEIGHT = 720
export const GRID_SIZE = 8
export const MIN_BLOCK_SIZE = 72

export const BlockRoleSchema = z.enum([
  'title',
  'subtitle',
  'body',
  'card',
  'metric',
  'chart',
  'note',
])
export const BlockTagSchema = z.enum(['h1', 'p', 'div'])

export const RectSchema = z.object({
  x: z.number(),
  y: z.number(),
  width: z.number(),
  height: z.number(),
})

export const SlideBlockSchema = RectSchema.extend({
  id: z.string(),
  role: BlockRoleSchema,
  tag: BlockTagSchema,
  text: z.string(),
  className: z.string(),
})

export const RetouchSlideSchema = z.object({
  id: z.string(),
  name: z.string(),
  accent: z.string(),
  blocks: z.array(SlideBlockSchema),
})

export const RetouchDeckSchema = z.object({
  slides: z.array(RetouchSlideSchema),
})

export type BlockRole = z.infer<typeof BlockRoleSchema>
export type BlockTag = z.infer<typeof BlockTagSchema>
export type Rect = z.infer<typeof RectSchema>
export type SlideBlock = z.infer<typeof SlideBlockSchema>
export type RetouchSlide = z.infer<typeof RetouchSlideSchema>
export type RetouchDeck = z.infer<typeof RetouchDeckSchema>

export type ResizeHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

export const RESIZE_HANDLES: ResizeHandle[] = [
  'nw',
  'n',
  'ne',
  'e',
  'se',
  's',
  'sw',
  'w',
]

export const SAMPLE_DECK: RetouchDeck = {
  slides: [
    {
      id: 'slide-1',
      name: 'Overview',
      accent: '#2563eb',
      blocks: [
        {
          id: 's1-title',
          role: 'title',
          tag: 'h1',
          text: 'Retention Review',
          className: 'block-title centered',
          x: 240,
          y: 158,
          width: 800,
          height: 84,
        },
        {
          id: 's1-subtitle',
          role: 'subtitle',
          tag: 'p',
          text: "Decide where this week's customer success effort should go.",
          className: 'block-subtitle centered',
          x: 280,
          y: 284,
          width: 720,
          height: 86,
        },
        {
          id: 's1-metric',
          role: 'metric',
          tag: 'div',
          text: 'Status\nNeeds attention',
          className: 'block-metric centered',
          x: 500,
          y: 418,
          width: 280,
          height: 112,
        },
        {
          id: 's1-note',
          role: 'note',
          tag: 'div',
          text: 'Next action: confirm owners and dates for the 12 at-risk accounts.',
          className: 'block-note centered',
          x: 280,
          y: 572,
          width: 720,
          height: 76,
        },
      ],
    },
    {
      id: 'slide-2',
      name: 'Agenda',
      accent: '#0f766e',
      blocks: [
        {
          id: 's2-title',
          role: 'title',
          tag: 'h1',
          text: "Today's Agenda",
          className: 'block-title compact centered',
          x: 300,
          y: 104,
          width: 680,
          height: 78,
        },
        {
          id: 's2-step-1',
          role: 'card',
          tag: 'div',
          text: '1\nRisk accounts',
          className: 'block-step centered',
          x: 168,
          y: 284,
          width: 260,
          height: 150,
        },
        {
          id: 's2-step-2',
          role: 'card',
          tag: 'div',
          text: '2\nOwner plan',
          className: 'block-step strong centered',
          x: 510,
          y: 284,
          width: 260,
          height: 150,
        },
        {
          id: 's2-step-3',
          role: 'card',
          tag: 'div',
          text: '3\nDecision needed',
          className: 'block-step centered',
          x: 852,
          y: 284,
          width: 260,
          height: 150,
        },
        {
          id: 's2-footer',
          role: 'body',
          tag: 'p',
          text: 'Goal: spend less time reporting and more time deciding.',
          className: 'block-footer centered',
          x: 300,
          y: 548,
          width: 680,
          height: 62,
        },
      ],
    },
    {
      id: 'slide-3',
      name: 'Decision',
      accent: '#4f46e5',
      blocks: [
        {
          id: 's3-title',
          role: 'title',
          tag: 'h1',
          text: 'Decision Needed',
          className: 'block-title wide centered',
          x: 320,
          y: 118,
          width: 640,
          height: 78,
        },
        {
          id: 's3-quote',
          role: 'body',
          tag: 'div',
          text: 'Approve a focused retention sprint for 12 accounts.',
          className: 'block-quote centered',
          x: 240,
          y: 260,
          width: 800,
          height: 130,
        },
        {
          id: 's3-card-1',
          role: 'metric',
          tag: 'div',
          text: 'Owner\nCS Lead',
          className: 'block-mini centered',
          x: 312,
          y: 452,
          width: 260,
          height: 96,
        },
        {
          id: 's3-card-2',
          role: 'metric',
          tag: 'div',
          text: 'Due\nFriday',
          className: 'block-mini secondary centered',
          x: 708,
          y: 452,
          width: 260,
          height: 96,
        },
        {
          id: 's3-note',
          role: 'note',
          tag: 'p',
          text: 'Decision: assign discount and onboarding support to the priority accounts.',
          className: 'block-note slim centered',
          x: 260,
          y: 584,
          width: 760,
          height: 76,
        },
      ],
    },
  ],
}

export const SAMPLE_SLIDES = SAMPLE_DECK.slides

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

export function blockPointer(slideIndex: number, blockIndex: number): Pointer {
  return buildPointer(['slides', slideIndex, 'blocks', blockIndex])
}

export function blockTextPointer(pointer: Pointer): Pointer {
  return appendSegment(pointer, 'text')
}

export function findSlideIndex(deck: RetouchDeck, slideId: string) {
  return deck.slides.findIndex((slide) => slide.id === slideId)
}

export function findBlockLocation(
  deck: RetouchDeck,
  slideId: string,
  blockId: string,
): BlockLocation | null {
  const slideIndex = findSlideIndex(deck, slideId)

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
  if (!pointer) {
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

  if (!slide || !block) {
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

export function getRect(block: SlideBlock): Rect {
  return {
    x: block.x,
    y: block.y,
    width: block.width,
    height: block.height,
  }
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

export function rectToStyle(rect: Rect) {
  return {
    left: `${(rect.x / SLIDE_WIDTH) * 100}%`,
    top: `${(rect.y / SLIDE_HEIGHT) * 100}%`,
    width: `${(rect.width / SLIDE_WIDTH) * 100}%`,
    height: `${(rect.height / SLIDE_HEIGHT) * 100}%`,
  }
}

export function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

export function snap(value: number) {
  return Math.round(value / GRID_SIZE) * GRID_SIZE
}

export function moveRect(rect: Rect, dx: number, dy: number): Rect {
  const x = clamp(snap(rect.x + dx), 0, SLIDE_WIDTH - rect.width)
  const y = clamp(snap(rect.y + dy), 0, SLIDE_HEIGHT - rect.height)

  return {
    ...rect,
    x,
    y,
  }
}

export function resizeRect(
  rect: Rect,
  handle: ResizeHandle,
  dx: number,
  dy: number,
): Rect {
  let left = rect.x
  let top = rect.y
  let right = rect.x + rect.width
  let bottom = rect.y + rect.height

  if (handle.includes('w')) {
    left = clamp(snap(left + dx), 0, right - MIN_BLOCK_SIZE)
  }

  if (handle.includes('e')) {
    right = clamp(snap(right + dx), left + MIN_BLOCK_SIZE, SLIDE_WIDTH)
  }

  if (handle.includes('n')) {
    top = clamp(snap(top + dy), 0, bottom - MIN_BLOCK_SIZE)
  }

  if (handle.includes('s')) {
    bottom = clamp(snap(bottom + dy), top + MIN_BLOCK_SIZE, SLIDE_HEIGHT)
  }

  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  }
}

export function rectEquals(a: Rect, b: Rect) {
  return (
    a.x === b.x &&
    a.y === b.y &&
    a.width === b.width &&
    a.height === b.height
  )
}

export function exportRetouchDeck(deck: RetouchDeck) {
  const css = [
    '<style>',
    ':root{font-family:Inter,ui-sans-serif,system-ui,sans-serif;color:#111827;background:#f3f4f6;}',
    '.deck{display:grid;gap:32px;padding:32px;}',
    '.slide{position:relative;width:1280px;height:720px;overflow:hidden;background:#fff;border:1px solid #d1d5db;}',
    '[data-block]{position:absolute;box-sizing:border-box;margin:0;overflow:hidden;white-space:pre-wrap;}',
    '.block-title{display:block;color:#111827;font:760 58px/1.04 Inter,system-ui,sans-serif;}',
    '.block-title.compact{font-size:54px;}',
    '.block-title.wide{font-size:52px;}',
    '.block-subtitle{display:block;color:#475569;font:450 24px/1.38 Inter,system-ui,sans-serif;}',
    '.block-metric{display:block;padding:20px;border:1px solid #dbe3ef;border-radius:8px;background:#f8fafc;color:#111827;font:650 25px/1.28 Inter,system-ui,sans-serif;}',
    '.block-chart{display:block;padding:20px;border:1px solid #dbe3ef;border-radius:8px;background:#fff;color:#111827;font:650 24px/1.36 Inter,system-ui,sans-serif;}',
    '.block-note{display:block;padding:18px 22px;border-top:2px solid var(--accent);background:#fff;color:#1f2937;font:620 24px/1.34 Inter,system-ui,sans-serif;}',
    '.block-note.slim{font-size:22px;}',
    '.block-step{display:block;padding:22px;border:1px solid #dbe3ef;border-radius:8px;background:#fff;color:#111827;font:650 27px/1.3 Inter,system-ui,sans-serif;}',
    '.block-step.strong{border-color:var(--accent);background:var(--accent);color:#fff;}',
    '.block-footer{display:block;color:#334155;font:650 26px/1.24 Inter,system-ui,sans-serif;}',
    '.block-quote{display:block;color:#111827;font:700 34px/1.18 Inter,system-ui,sans-serif;}',
    '.block-mini{display:block;padding:18px;border:1px solid #dbe3ef;border-radius:8px;background:#f8fafc;color:#111827;font:650 24px/1.24 Inter,system-ui,sans-serif;}',
    '.block-mini.secondary{background:#fff;color:#1f2937;}',
    '.centered{text-align:center;}',
    '</style>',
  ].join('\n')

  const html = deck.slides
    .map((slide) => {
      const blocks = slide.blocks
        .map((block) => {
          const text = escapeHtml(block.text)
          const style = [
            `left:${block.x}px`,
            `top:${block.y}px`,
            `width:${block.width}px`,
            `height:${block.height}px`,
          ].join(';')

          return `    <${block.tag} data-block="${block.id}" data-role="${block.role}" class="${block.className}" style="${style}">${text}</${block.tag}>`
        })
        .join('\n')

      return `  <section data-slide="${slide.id}" class="slide" style="--accent:${slide.accent}">\n${blocks}\n  </section>`
    })
    .join('\n')

  return `${css}\n\n<main class="deck">\n${html}\n</main>\n`
}

function escapeHtml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}
