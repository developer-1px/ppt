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
      name: 'Pipeline',
      accent: '#2563eb',
      blocks: [
        {
          id: 's1-title',
          role: 'title',
          tag: 'h1',
          text: 'Q3 Pipeline Review',
          className: 'block-title',
          x: 76,
          y: 66,
          width: 590,
          height: 118,
        },
        {
          id: 's1-subtitle',
          role: 'subtitle',
          tag: 'p',
          text: 'Enterprise expansion is ahead of plan, while mid-market renewals need closer follow-up before month end.',
          className: 'block-subtitle',
          x: 80,
          y: 206,
          width: 560,
          height: 140,
        },
        {
          id: 's1-metric',
          role: 'metric',
          tag: 'div',
          text: '42%\nenterprise growth',
          className: 'block-metric',
          x: 746,
          y: 78,
          width: 360,
          height: 172,
        },
        {
          id: 's1-chart',
          role: 'chart',
          tag: 'div',
          text: 'Inbound\nExpansion\nRenewal\nPartner',
          className: 'block-chart',
          x: 724,
          y: 304,
          width: 398,
          height: 274,
        },
        {
          id: 's1-note',
          role: 'note',
          tag: 'div',
          text: 'Focus the next two weeks on renewal risk and partner-sourced opportunities.',
          className: 'block-note',
          x: 80,
          y: 468,
          width: 506,
          height: 120,
        },
      ],
    },
    {
      id: 'slide-2',
      name: 'Plan',
      accent: '#0f766e',
      blocks: [
        {
          id: 's2-title',
          role: 'title',
          tag: 'h1',
          text: 'Operating Plan',
          className: 'block-title compact',
          x: 82,
          y: 62,
          width: 460,
          height: 96,
        },
        {
          id: 's2-step-1',
          role: 'card',
          tag: 'div',
          text: '1\nPrioritize enterprise expansion.',
          className: 'block-step',
          x: 96,
          y: 230,
          width: 290,
          height: 240,
        },
        {
          id: 's2-step-2',
          role: 'card',
          tag: 'div',
          text: '2\nRecover renewal risk early.',
          className: 'block-step strong',
          x: 494,
          y: 188,
          width: 290,
          height: 282,
        },
        {
          id: 's2-step-3',
          role: 'card',
          tag: 'div',
          text: '3\nTighten partner handoffs.',
          className: 'block-step',
          x: 892,
          y: 230,
          width: 290,
          height: 240,
        },
        {
          id: 's2-footer',
          role: 'body',
          tag: 'p',
          text: 'Weekly review keeps the forecast stable.',
          className: 'block-footer',
          x: 344,
          y: 562,
          width: 592,
          height: 62,
        },
      ],
    },
    {
      id: 'slide-3',
      name: 'Board',
      accent: '#7c3aed',
      blocks: [
        {
          id: 's3-title',
          role: 'title',
          tag: 'h1',
          text: 'Board Update',
          className: 'block-title wide',
          x: 88,
          y: 72,
          width: 780,
          height: 112,
        },
        {
          id: 's3-quote',
          role: 'body',
          tag: 'div',
          text: 'The base forecast remains intact, but two enterprise renewals need executive attention before close.',
          className: 'block-quote',
          x: 100,
          y: 244,
          width: 620,
          height: 208,
        },
        {
          id: 's3-card-1',
          role: 'metric',
          tag: 'div',
          text: 'Revenue\n+18% QoQ',
          className: 'block-mini',
          x: 820,
          y: 226,
          width: 260,
          height: 130,
        },
        {
          id: 's3-card-2',
          role: 'metric',
          tag: 'div',
          text: 'Risk\nrenewals',
          className: 'block-mini secondary',
          x: 820,
          y: 386,
          width: 260,
          height: 130,
        },
        {
          id: 's3-note',
          role: 'note',
          tag: 'p',
          text: 'Decision needed: approve the partner incentive pilot.',
          className: 'block-note slim',
          x: 102,
          y: 536,
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
    '.block-title{display:flex;align-items:center;color:#111827;font:760 58px/1.04 Inter,system-ui,sans-serif;}',
    '.block-title.compact{font-size:54px;}',
    '.block-title.wide{font-size:52px;}',
    '.block-subtitle{display:flex;align-items:center;color:#475569;font:450 24px/1.38 Inter,system-ui,sans-serif;}',
    '.block-metric{display:flex;align-items:center;padding:24px;border:1px solid #bfdbfe;border-radius:8px;background:#eff6ff;color:#172554;font:650 25px/1.28 Inter,system-ui,sans-serif;}',
    '.block-chart{display:flex;align-items:center;padding:26px;border-radius:8px;background:linear-gradient(90deg,var(--accent) 0 18%,transparent 18% 100%) 0 20%/100% 18% no-repeat,linear-gradient(90deg,#38bdf8 0 42%,transparent 42% 100%) 0 44%/100% 18% no-repeat,linear-gradient(90deg,#f59e0b 0 68%,transparent 68% 100%) 0 68%/100% 18% no-repeat,linear-gradient(90deg,#22c55e 0 86%,transparent 86% 100%) 0 92%/100% 18% no-repeat,#111827;color:#f8fafc;font:650 28px/1.45 Inter,system-ui,sans-serif;}',
    '.block-note{display:flex;align-items:center;padding:22px;border-left:6px solid var(--accent);background:#f8fafc;color:#1f2937;font:620 24px/1.34 Inter,system-ui,sans-serif;}',
    '.block-note.slim{font-size:22px;}',
    '.block-step{display:flex;align-items:center;padding:28px;border:1px solid #ccfbf1;border-radius:8px;background:#f0fdfa;color:#134e4a;font:650 27px/1.3 Inter,system-ui,sans-serif;}',
    '.block-step.strong{border-color:#0f766e;background:#0f766e;color:#fff;}',
    '.block-footer{display:flex;align-items:center;justify-content:center;color:#334155;font:650 26px/1.24 Inter,system-ui,sans-serif;text-align:center;}',
    '.block-quote{display:flex;align-items:center;padding:28px;border-radius:8px;background:#f5f3ff;color:#2e1065;font:650 31px/1.3 Inter,system-ui,sans-serif;}',
    '.block-mini{display:flex;align-items:center;padding:22px;border-radius:8px;background:#ede9fe;color:#2e1065;font:720 27px/1.22 Inter,system-ui,sans-serif;}',
    '.block-mini.secondary{background:#f3f4f6;color:#1f2937;}',
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
