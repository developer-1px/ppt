import {
  clamp,
  type Point,
  type Viewport,
} from 'canvas/core'
import {
  getCanvasTextPasteSourcesFromDataTransfer,
} from 'canvas/app/text-paste-import'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTParagraph,
  type PPTRun,
  type PPTTextBody,
  type PPTTextBox,
} from '../pptModel'

type CanvasTextPasteTextItem = {
  fontSize?: number
  h: number
  id: string
  text: string
  type: 'text'
  w: number
  x: number
  y: number
}

export type PPTTextPasteImportResult = {
  boldRunCount?: number
  bulletParagraphCount?: number
  format: 'text-html-rich' | 'text-plain'
  importerId: string
  italicRunCount?: number
  item: PPTTextBox
  linkRunCount?: number
  underlineRunCount?: number
}
export type PPTRichTextPasteSource = {
  boldRunCount: number
  bulletParagraphCount: number
  linkRunCount: number
  text: string
  textBody: PPTTextBody
  underlineRunCount: number
}

type PPTTextPasteImporter = {
  createItems: (context: {
    createId: (prefix: string) => string
    position: Point
    text: string
    viewport: Viewport
  }) => CanvasTextPasteTextItem[]
  id: string
}

const PPT_TEXT_PASTE_WIDTH = 460
const PPT_TEXT_PASTE_LINE_HEIGHT = 38
const PPT_TEXT_PASTE_MIN_HEIGHT = 92
const PPT_TEXT_PASTE_MAX_HEIGHT = 320
export const getPPTTextPasteSourcesFromDataTransfer =
  getCanvasTextPasteSourcesFromDataTransfer

export function getPPTRichTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTRichTextPasteSource | null {
  if (!dataTransfer) {
    return null
  }

  return getPPTRichTextPasteSourceFromHTML(dataTransfer.getData('text/html'))
}

export function createPPTTextPasteElement({
  createId,
  position,
  text,
  viewport,
}: {
  createId: (prefix: string) => string
  position: Point
  text: string
  viewport: Viewport
}): PPTTextPasteImportResult | null {
  const trimmedText = text.trim()

  if (!trimmedText) {
    return null
  }

  let items: CanvasTextPasteTextItem[] | null

  try {
    items = PPT_TEXT_PASTE_IMPORTER.createItems({
      createId,
      position,
      text: trimmedText,
      viewport,
    })
  } catch {
    items = null
  }

  const canvasItem = items?.find(isCanvasTextPasteTextItem)

  if (!canvasItem) {
    return null
  }

  return {
    format: 'text-plain',
    importerId: PPT_TEXT_PASTE_IMPORTER.id,
    item: {
      geometry: {
        h: canvasItem.h,
        w: canvasItem.w,
        x: clamp(canvasItem.x, 0, PPT_SLIDE_WIDTH - canvasItem.w),
        y: clamp(canvasItem.y, 0, PPT_SLIDE_HEIGHT - canvasItem.h),
      },
      id: canvasItem.id,
      kind: 'textBox',
      name: 'Text',
      style: {
        color: '#111827',
        fontSize: canvasItem.fontSize ?? 30,
        fontWeight: 'semibold',
      },
      textBody: createPPTTextBody(canvasItem.text),
    },
  }
}

export function createPPTRichTextPasteElement({
  createId,
  position,
  source,
  viewport,
}: {
  createId: (prefix: string) => string
  position: Point
  source: PPTRichTextPasteSource
  viewport: Viewport
}): PPTTextPasteImportResult | null {
  const result = createPPTTextPasteElement({
    createId,
    position,
    text: source.text,
    viewport,
  })

  if (!result) {
    return null
  }

  return {
    boldRunCount: source.boldRunCount,
    bulletParagraphCount: source.bulletParagraphCount,
    format: 'text-html-rich',
    importerId: 'ppt-rich-html-text',
    italicRunCount: getPPTTextBodyRunCount(source.textBody, 'italic'),
    item: {
      ...result.item,
      name: 'Rich Text',
      textBody: source.textBody,
    },
    linkRunCount: source.linkRunCount,
    underlineRunCount: source.underlineRunCount,
  }
}

const PPT_TEXT_PASTE_IMPORTER: PPTTextPasteImporter = {
  id: 'ppt-plain-text',
  createItems: ({ createId, position, text }: {
    createId: (prefix: string) => string
    position: Point
    text: string
  }) => {
    const lineCount = Math.max(1, text.split('\n').length)
    const height = clamp(
      lineCount * PPT_TEXT_PASTE_LINE_HEIGHT + 32,
      PPT_TEXT_PASTE_MIN_HEIGHT,
      PPT_TEXT_PASTE_MAX_HEIGHT,
    )

    return [{
      fontSize: 30,
      h: height,
      id: createId('text'),
      text,
      type: 'text',
      w: PPT_TEXT_PASTE_WIDTH,
      x: position.x - PPT_TEXT_PASTE_WIDTH / 2,
      y: position.y - height / 2,
    } as const]
  },
}

function isCanvasTextPasteTextItem(
  value: unknown,
): value is CanvasTextPasteTextItem {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as CanvasTextPasteTextItem).type === 'text' &&
    typeof (value as CanvasTextPasteTextItem).id === 'string' &&
    typeof (value as CanvasTextPasteTextItem).text === 'string' &&
    typeof (value as CanvasTextPasteTextItem).x === 'number' &&
    typeof (value as CanvasTextPasteTextItem).y === 'number' &&
    typeof (value as CanvasTextPasteTextItem).w === 'number' &&
    typeof (value as CanvasTextPasteTextItem).h === 'number'
  )
}

function getPPTRichTextPasteSourceFromHTML(value: string) {
  if (!value || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(value, 'text/html')

  if (doc.body.querySelector('table, img, svg, video, audio, iframe')) {
    return null
  }

  doc.body.querySelectorAll('script, style, noscript').forEach((node) => node.remove())

  const paragraphs = getPPTRichTextParagraphs(doc.body)
  const textBody = normalizePPTRichTextBody({ paragraphs })
  const text = textBody.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n')

  if (!text.trim() || !hasPPTRichTextFormatting(textBody)) {
    return null
  }

  return {
    boldRunCount: getPPTTextBodyRunCount(textBody, 'bold'),
    bulletParagraphCount: textBody.paragraphs.filter((paragraph) => paragraph.bullet === 'bullet').length,
    linkRunCount: getPPTTextBodyRunCount(textBody, 'color'),
    text,
    textBody,
    underlineRunCount: getPPTTextBodyRunCount(textBody, 'underline'),
  }
}

function getPPTRichTextParagraphs(root: HTMLElement): PPTParagraph[] {
  const blockNodes = Array.from(root.children).filter(isPPTRichTextBlock)

  if (blockNodes.length === 0) {
    const runs = getPPTRichTextRuns(root, {})

    return runs.length > 0 ? [{ runs }] : []
  }

  return blockNodes.flatMap((node) => getPPTRichTextParagraphFromBlock(node))
}

function getPPTRichTextParagraphFromBlock(node: Element): PPTParagraph[] {
  if (node.tagName.toLowerCase() === 'br') {
    return []
  }

  const nestedBlocks = Array.from(node.children).filter((child) =>
    isPPTRichTextBlock(child) && child.tagName.toLowerCase() !== 'br',
  )

  if (
    nestedBlocks.length > 0 &&
    !['li', 'p'].includes(node.tagName.toLowerCase())
  ) {
    return nestedBlocks.flatMap((child) => getPPTRichTextParagraphFromBlock(child))
  }

  const runs = getPPTRichTextRuns(node, {})
  const bullet = node.tagName.toLowerCase() === 'li' ? 'bullet' : undefined

  return runs.length > 0
    ? [{
        ...(bullet ? { bullet } : {}),
        runs,
      }]
    : []
}

function getPPTRichTextRuns(
  node: Node,
  style: Partial<PPTRun>,
): PPTRun[] {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent?.replace(/\s+/g, ' ') ?? ''

    return text.trim().length > 0
      ? [{
          ...style,
          text,
        }]
      : []
  }

  if (!(node instanceof Element)) {
    return []
  }

  const tagName = node.tagName.toLowerCase()

  if (tagName === 'br') {
    return [{ text: '\n' }]
  }

  const nextStyle = {
    ...style,
    ...(['b', 'strong'].includes(tagName) ? { bold: true } : {}),
    ...(['em', 'i'].includes(tagName) ? { italic: true } : {}),
    ...(['a', 'u'].includes(tagName) ? { underline: true } : {}),
    ...(tagName === 'a' ? { color: '#2563eb' } : {}),
  }

  return Array.from(node.childNodes).flatMap((child) =>
    getPPTRichTextRuns(child, nextStyle))
}

function normalizePPTRichTextBody(body: PPTTextBody): PPTTextBody {
  return {
    paragraphs: body.paragraphs.flatMap((paragraph) => {
      const splitParagraphs: PPTParagraph[] = []
      let runs: PPTRun[] = []

      for (const run of paragraph.runs) {
        const chunks = run.text.split('\n')

        chunks.forEach((chunk, index) => {
          if (index > 0) {
            if (runs.length > 0) {
              splitParagraphs.push({
                ...(paragraph.bullet ? { bullet: paragraph.bullet } : {}),
                runs,
              })
            }
            runs = []
          }

          if (chunk.trim()) {
            runs.push({
              ...run,
              text: chunk,
            })
          }
        })
      }

      if (runs.length > 0) {
        splitParagraphs.push({
          ...(paragraph.bullet ? { bullet: paragraph.bullet } : {}),
          runs,
        })
      }

      return splitParagraphs
    }),
  }
}

function hasPPTRichTextFormatting(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) =>
    paragraph.bullet === 'bullet' ||
    paragraph.runs.some((run) =>
      run.bold === true ||
      run.italic === true ||
      run.underline === true ||
      Boolean(run.color)))
}

function getPPTTextBodyRunCount(
  body: PPTTextBody,
  field: keyof Pick<PPTRun, 'bold' | 'color' | 'italic' | 'underline'>,
) {
  return body.paragraphs.reduce((count, paragraph) =>
    count + paragraph.runs.filter((run) => Boolean(run[field])).length, 0)
}

function isPPTRichTextBlock(node: Element) {
  return [
    'article',
    'blockquote',
    'br',
    'div',
    'h1',
    'h2',
    'h3',
    'h4',
    'h5',
    'h6',
    'li',
    'main',
    'ol',
    'p',
    'section',
    'ul',
  ].includes(node.tagName.toLowerCase())
}
