export const SLIDE_WIDTH = 1280
export const SLIDE_HEIGHT = 720
export const GRID_SIZE = 8
export const MIN_BLOCK_SIZE = 72

export type BlockRole =
  | 'title'
  | 'subtitle'
  | 'body'
  | 'card'
  | 'metric'
  | 'chart'
  | 'note'

export type BlockTag = 'h1' | 'p' | 'div'

export type Rect = {
  x: number
  y: number
  width: number
  height: number
}

export type SlideBlock = Rect & {
  id: string
  role: BlockRole
  tag: BlockTag
  text: string
  className: string
}

export type RetouchSlide = {
  id: string
  name: string
  accent: string
  blocks: SlideBlock[]
}

export type TextPatch = {
  blockId: string
  text: string
}

export type LayoutPatch = {
  blockId: string
  rect: Rect
}

export type SlideEdits = {
  textPatches: Record<string, string>
  layoutPatches: Record<string, Rect>
}

export type DeckEdits = Record<string, SlideEdits>

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

export const SAMPLE_SLIDES: RetouchSlide[] = [
  {
    id: 'slide-1',
    name: 'Market',
    accent: '#2563eb',
    blocks: [
      {
        id: 's1-title',
        role: 'title',
        tag: 'h1',
        text: 'AI Workspace Adoption',
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
        text: 'Teams keep the generated draft, then spend the final pass on wording and layout confidence.',
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
        text: '64%\nprefer retouching',
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
        text: 'Draft\nReview\nRetouch\nShare',
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
        text: 'The winning workflow separates text edits from layout edits.',
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
    name: 'Flow',
    accent: '#0f766e',
    blocks: [
      {
        id: 's2-title',
        role: 'title',
        tag: 'h1',
        text: 'Retouch Flow',
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
        text: '1\nAI creates the HTML/CSS slide.',
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
        text: '2\nText mode changes only the words.',
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
        text: '3\nLayout mode adjusts the final fit.',
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
        text: 'Each mode removes the other kind of risk.',
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
    name: 'Close',
    accent: '#7c3aed',
    blocks: [
      {
        id: 's3-title',
        role: 'title',
        tag: 'h1',
        text: 'Final Pass, Not Full Design',
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
        text: 'The tool should feel like the last five minutes before sharing: direct, reversible, and hard to break.',
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
        text: 'Text\nwords only',
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
        text: 'Layout\ndrag and resize',
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
        text: 'No CSS controls. No numeric typography. No mixed-mode editing.',
        className: 'block-note slim',
        x: 102,
        y: 536,
        width: 760,
        height: 76,
      },
    ],
  },
]

export function createInitialDeckEdits(slides: RetouchSlide[]): DeckEdits {
  return Object.fromEntries(
    slides.map((slide) => [
      slide.id,
      {
        textPatches: {},
        layoutPatches: {},
      },
    ]),
  )
}

export function getText(block: SlideBlock, slideEdits: SlideEdits) {
  return slideEdits.textPatches[block.id] ?? block.text
}

export function getRect(block: SlideBlock, slideEdits: SlideEdits): Rect {
  return slideEdits.layoutPatches[block.id] ?? {
    x: block.x,
    y: block.y,
    width: block.width,
    height: block.height,
  }
}

export function setTextPatch(
  edits: DeckEdits,
  slideId: string,
  blockId: string,
  text: string,
): DeckEdits {
  const slideEdits = edits[slideId]
  return {
    ...edits,
    [slideId]: {
      ...slideEdits,
      textPatches: {
        ...slideEdits.textPatches,
        [blockId]: text,
      },
    },
  }
}

export function setLayoutPatch(
  edits: DeckEdits,
  slideId: string,
  blockId: string,
  rect: Rect,
): DeckEdits {
  const slideEdits = edits[slideId]
  return {
    ...edits,
    [slideId]: {
      ...slideEdits,
      layoutPatches: {
        ...slideEdits.layoutPatches,
        [blockId]: rect,
      },
    },
  }
}

export function resetLayoutPatch(
  edits: DeckEdits,
  slideId: string,
  blockId: string,
): DeckEdits {
  const slideEdits = edits[slideId]
  const layoutPatches = { ...slideEdits.layoutPatches }
  delete layoutPatches[blockId]

  return {
    ...edits,
    [slideId]: {
      ...slideEdits,
      layoutPatches,
    },
  }
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

export function exportRetouchDeck(slides: RetouchSlide[], edits: DeckEdits) {
  const css = [
    '<style>',
    ':root{font-family:Inter,ui-sans-serif,system-ui,sans-serif;color:#111827;background:#f3f4f6;}',
    '.deck{display:grid;gap:32px;padding:32px;}',
    '.slide{position:relative;width:1280px;height:720px;overflow:hidden;background:#fff;border:1px solid #d1d5db;}',
    '[data-block]{position:absolute;box-sizing:border-box;white-space:pre-wrap;}',
    '.block-title{font:760 58px/1.04 Inter,system-ui,sans-serif;}',
    '.block-title.compact{font-size:54px;}',
    '.block-title.wide{font-size:52px;}',
    '.block-subtitle{font:450 24px/1.38 Inter,system-ui,sans-serif;color:#475569;}',
    '.block-metric{padding:24px;border:1px solid #bfdbfe;border-radius:8px;background:#eff6ff;font:650 25px/1.28 Inter,system-ui,sans-serif;}',
    '.block-chart{padding:26px;border-radius:8px;background:#111827;color:#f8fafc;font:650 28px/1.45 Inter,system-ui,sans-serif;}',
    '.block-note{padding:22px;border-left:6px solid #2563eb;background:#f8fafc;font:620 24px/1.34 Inter,system-ui,sans-serif;}',
    '.block-note.slim{font-size:22px;}',
    '.block-step{padding:28px;border:1px solid #ccfbf1;border-radius:8px;background:#f0fdfa;font:650 27px/1.3 Inter,system-ui,sans-serif;}',
    '.block-step.strong{background:#0f766e;color:white;}',
    '.block-footer{font:650 26px/1.24 Inter,system-ui,sans-serif;text-align:center;color:#334155;}',
    '.block-quote{padding:28px;border-radius:8px;background:#f5f3ff;font:650 31px/1.3 Inter,system-ui,sans-serif;}',
    '.block-mini{padding:22px;border-radius:8px;background:#ede9fe;font:720 27px/1.22 Inter,system-ui,sans-serif;}',
    '.block-mini.secondary{background:#f3f4f6;}',
    '</style>',
  ].join('\n')

  const html = slides
    .map((slide) => {
      const slideEdits = edits[slide.id]
      const blocks = slide.blocks
        .map((block) => {
          const text = escapeHtml(getText(block, slideEdits))
          const rect = getRect(block, slideEdits)
          const style = [
            `left:${rect.x}px`,
            `top:${rect.y}px`,
            `width:${rect.width}px`,
            `height:${rect.height}px`,
          ].join(';')

          return `    <${block.tag} data-block="${block.id}" data-role="${block.role}" class="${block.className}" style="${style}">${text}</${block.tag}>`
        })
        .join('\n')

      return `  <section data-slide="${slide.id}" class="slide">\n${blocks}\n  </section>`
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
