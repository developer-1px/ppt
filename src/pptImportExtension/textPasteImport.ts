import {
  clampPPTCanvasBoundsToFrame,
  clampPPTCanvasValue,
  type Point,
  type Viewport,
} from '../pptCanvasCoreAdapter'
import {
  createPPTCanvasTextPasteItems,
  getPPTCanvasRichTextPasteSourceFromDataTransfer,
  getPPTCanvasTextPasteSourcesFromDataTransfer,
  PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL,
  type PPTCanvasRichTextPasteSource,
} from '../pptCanvasAppAffordanceAdapter'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTParagraph,
  type PPTRun,
  type PPTTextBody,
  type PPTTextBox,
} from '../pptModel'

type PPTTextPasteTextItem = {
  fontSize?: number
  h: number
  id: string
  text: string
  type: 'text'
  w: number
  x: number
  y: number
}

type PPTTextPasteImporter =
  Parameters<typeof createPPTCanvasTextPasteItems>[0]['importers'][number]

export type PPTTextPasteImportResult = {
  boldRunCount?: number
  bulletParagraphCount?: number
  format: 'text-html-rich' | 'text-markdown-rich' | 'text-plain'
  importerId: string
  italicRunCount?: number
  item: PPTTextBox
  linkRunCount?: number
  numberedParagraphCount?: number
  underlineRunCount?: number
}
export type PPTRichTextPasteSource = {
  boldRunCount: number
  bulletParagraphCount: number
  format?: 'text-html-rich' | 'text-markdown-rich'
  importerId?: string
  linkRunCount: number
  name?: string
  numberedParagraphCount: number
  text: string
  textBody: PPTTextBody
  underlineRunCount: number
}

const PPT_TEXT_PASTE_WIDTH = 460
const PPT_TEXT_PASTE_LINE_HEIGHT = 38
const PPT_TEXT_PASTE_MIN_HEIGHT = 92
const PPT_TEXT_PASTE_MAX_HEIGHT = 320
const PPT_HTML_RICH_TEXT_BLOCK_TAGS = new Set([
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
])
export const PPT_TEXT_PASTE_IMPORT_MODEL = PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL

export const getPPTTextPasteSourcesFromDataTransfer =
  getPPTCanvasTextPasteSourcesFromDataTransfer

export function getPPTRichTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTRichTextPasteSource | null {
  const source = getPPTCanvasRichTextPasteSourceFromDataTransfer(dataTransfer)

  return source
    ? createPPTRichTextPasteSource(source, {
        html: dataTransfer?.getData('text/html') ?? '',
      })
    : null
}

export function getPPTMarkdownTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTRichTextPasteSource | null {
  if (!dataTransfer) {
    return null
  }

  return getPPTMarkdownTextPasteSourceFromText(
    dataTransfer.getData('text/markdown') ||
      dataTransfer.getData('text/plain'),
  )
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
  const result = createPPTCanvasTextPasteItems({
    createId,
    importers: [PPT_TEXT_PASTE_IMPORTER],
    position,
    text,
    viewport,
  })

  if (!result) {
    return null
  }

  const textItem = result.items.find(isPPTTextPasteTextItem)

  if (!textItem) {
    return null
  }

  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: textItem.h,
      w: textItem.w,
      x: textItem.x,
      y: textItem.y,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    format: 'text-plain',
    importerId: result.importerId,
    item: {
      geometry,
      id: textItem.id,
      kind: 'textBox',
      name: 'Text',
      style: {
        color: '#111827',
        fontSize: textItem.fontSize ?? 30,
        fontWeight: 'semibold',
      },
      textBody: createPPTTextBody(textItem.text),
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
    format: source.format ?? 'text-html-rich',
    importerId: source.importerId ?? 'ppt-rich-html-text',
    italicRunCount: getPPTTextBodyRunCount(source.textBody, 'italic'),
    item: {
      ...result.item,
      name: source.name ?? 'Rich Text',
      textBody: source.textBody,
    },
    linkRunCount: source.linkRunCount,
    numberedParagraphCount: source.numberedParagraphCount,
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
    const height = clampPPTCanvasValue(
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

function isPPTTextPasteTextItem(
  value: unknown,
): value is PPTTextPasteTextItem {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as PPTTextPasteTextItem).type === 'text' &&
    typeof (value as PPTTextPasteTextItem).id === 'string' &&
    typeof (value as PPTTextPasteTextItem).text === 'string' &&
    typeof (value as PPTTextPasteTextItem).x === 'number' &&
    typeof (value as PPTTextPasteTextItem).y === 'number' &&
    typeof (value as PPTTextPasteTextItem).w === 'number' &&
    typeof (value as PPTTextPasteTextItem).h === 'number'
  )
}

function createPPTRichTextPasteSource(
  source: PPTCanvasRichTextPasteSource,
  options: { html?: string } = {},
): PPTRichTextPasteSource {
  const listKinds = getPPTHTMLRichTextParagraphListKinds(options.html ?? '')
  const textBody: PPTTextBody = {
    paragraphs: source.paragraphs.map((paragraph, index): PPTParagraph => {
      const bullet = listKinds[index] ?? paragraph.bullet

      return {
        ...(bullet ? { bullet } : {}),
        runs: paragraph.runs.map((run): PPTRun => ({
          ...(run.bold === undefined ? {} : { bold: run.bold }),
          ...(run.color || run.link ? { color: run.color ?? '#2563eb' } : {}),
          ...(run.italic === undefined ? {} : { italic: run.italic }),
          ...(run.underline === undefined ? {} : { underline: run.underline }),
          text: run.text,
        })),
      }
    }),
  }

  return {
    boldRunCount: getPPTTextBodyRunCount(textBody, 'bold'),
    bulletParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'bullet').length,
    linkRunCount: source.paragraphs.reduce((count, paragraph) =>
      count + paragraph.runs.filter((run) => Boolean(run.link)).length, 0),
    numberedParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'numbered').length,
    text: source.text,
    textBody,
    underlineRunCount: getPPTTextBodyRunCount(textBody, 'underline'),
  }
}

function getPPTMarkdownTextPasteSourceFromText(
  text: string,
): PPTRichTextPasteSource | null {
  const paragraphs = getPPTMarkdownTextPasteParagraphs(text)

  if (paragraphs.length === 0 || !hasPPTMarkdownFormatting(text, paragraphs)) {
    return null
  }

  const textBody = { paragraphs }

  return {
    boldRunCount: getPPTTextBodyRunCount(textBody, 'bold'),
    bulletParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'bullet').length,
    format: 'text-markdown-rich',
    importerId: 'ppt-rich-markdown-text',
    linkRunCount: getPPTTextBodyRunCount(textBody, 'color'),
    name: 'Markdown Text',
    numberedParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'numbered').length,
    text: textBody.paragraphs
      .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
      .join('\n'),
    textBody,
    underlineRunCount: getPPTTextBodyRunCount(textBody, 'underline'),
  }
}

function getPPTMarkdownTextPasteParagraphs(text: string): PPTParagraph[] {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map(parsePPTMarkdownTextPasteParagraph)
    .filter((paragraph): paragraph is PPTParagraph => paragraph !== null)
}

function parsePPTMarkdownTextPasteParagraph(line: string): PPTParagraph | null {
  const heading = line.match(/^#{1,6}\s+(.+?)\s*#*$/)
  const bullet = line.match(/^[-*+]\s+(.+)$/)
  const numbered = line.match(/^\d+[.)]\s+(.+)$/)
  const text = heading?.[1] ?? bullet?.[1] ?? numbered?.[1] ?? line
  const runs = parsePPTMarkdownTextRuns(text)

  if (runs.length === 0) {
    return null
  }

  return {
    ...(bullet ? { bullet: 'bullet' as const } : {}),
    ...(numbered ? { bullet: 'numbered' as const } : {}),
    runs: heading
      ? runs.map((run) => ({ ...run, bold: true }))
      : runs,
  }
}

function parsePPTMarkdownTextRuns(text: string): PPTRun[] {
  const runs: PPTRun[] = []
  const pattern =
    /(\[([^\]]+)]\(([^)]+)\)|`([^`]+)`|\*\*([^*]+)\*\*|__([^_]+)__|\*([^*]+)\*|_([^_]+)_)/g
  let index = 0

  for (const match of text.matchAll(pattern)) {
    if (match.index === undefined) {
      continue
    }

    if (match.index > index) {
      runs.push({ text: unescapePPTMarkdownText(text.slice(index, match.index)) })
    }

    if (match[2] !== undefined) {
      runs.push({
        color: '#2563eb',
        text: unescapePPTMarkdownText(match[2]),
        underline: true,
      })
    } else if (match[4] !== undefined) {
      runs.push({ text: unescapePPTMarkdownText(match[4]) })
    } else if (match[5] !== undefined || match[6] !== undefined) {
      runs.push({
        bold: true,
        text: unescapePPTMarkdownText(match[5] ?? match[6] ?? ''),
      })
    } else if (match[7] !== undefined || match[8] !== undefined) {
      runs.push({
        italic: true,
        text: unescapePPTMarkdownText(match[7] ?? match[8] ?? ''),
      })
    }

    index = match.index + match[0].length
  }

  if (index < text.length) {
    runs.push({ text: unescapePPTMarkdownText(text.slice(index)) })
  }

  return runs.filter((run) => run.text.length > 0)
}

function unescapePPTMarkdownText(text: string) {
  return text.replace(/\\([\\`*_[\]()#+.!|-])/g, '$1')
}

function hasPPTMarkdownFormatting(
  text: string,
  paragraphs: readonly PPTParagraph[],
) {
  return /^#{1,6}\s+/m.test(text) ||
    /^(?:[-*+]|\d+[.)])\s+/m.test(text) ||
    /\[([^\]]+)]\(([^)]+)\)/.test(text) ||
    /(`[^`]+`|\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_)/.test(text) ||
    paragraphs.some((paragraph) =>
      paragraph.bullet === 'bullet' ||
      paragraph.bullet === 'numbered' ||
      paragraph.runs.some((run) =>
        run.bold === true ||
        run.italic === true ||
        run.underline === true ||
        Boolean(run.color)))
}

function getPPTHTMLRichTextParagraphListKinds(
  html: string,
): Array<PPTParagraph['bullet'] | undefined> {
  if (!html || typeof DOMParser === 'undefined') {
    return []
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')

  doc.body.querySelectorAll('script, style, noscript')
    .forEach((node) => node.remove())

  return getPPTHTMLRichTextParagraphListKindsFromElement(doc.body)
}

function getPPTHTMLRichTextParagraphListKindsFromElement(
  root: Element,
): Array<PPTParagraph['bullet'] | undefined> {
  const blockNodes = Array.from(root.children).filter(isPPTHTMLRichTextBlock)

  if (blockNodes.length === 0) {
    return root.textContent?.trim() ? [undefined] : []
  }

  return blockNodes.flatMap((node) =>
    getPPTHTMLRichTextParagraphListKindFromBlock(node))
}

function getPPTHTMLRichTextParagraphListKindFromBlock(
  node: Element,
): Array<PPTParagraph['bullet'] | undefined> {
  const tagName = node.tagName.toLowerCase()

  if (tagName === 'br') {
    return []
  }

  const nestedBlocks = Array.from(node.children).filter((child) =>
    isPPTHTMLRichTextBlock(child) && child.tagName.toLowerCase() !== 'br')

  if (nestedBlocks.length > 0 && !['li', 'p'].includes(tagName)) {
    return nestedBlocks.flatMap((child) =>
      getPPTHTMLRichTextParagraphListKindFromBlock(child))
  }

  if (!node.textContent?.trim()) {
    return []
  }

  if (tagName !== 'li') {
    return [undefined]
  }

  return [node.parentElement?.tagName.toLowerCase() === 'ol'
    ? 'numbered'
    : 'bullet']
}

function isPPTHTMLRichTextBlock(node: Element) {
  return PPT_HTML_RICH_TEXT_BLOCK_TAGS.has(node.tagName.toLowerCase())
}

function getPPTTextBodyRunCount(
  body: PPTTextBody,
  field: keyof Pick<PPTRun, 'bold' | 'color' | 'italic' | 'underline'>,
) {
  return body.paragraphs.reduce((count, paragraph) =>
    count + paragraph.runs.filter((run) => Boolean(run[field])).length, 0)
}
