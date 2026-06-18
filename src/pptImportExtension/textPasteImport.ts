import {
  clampPPTCanvasBoundsToFrame,
  clampPPTCanvasValue,
  type Point,
  type Viewport,
} from '../pptCanvasCoreAdapter'
import {
  createPPTCanvasTextPasteItems,
  getPPTCanvasTextPasteSourceCandidatesFromDataTransfer,
  getPPTCanvasTextPasteSourceText,
  getPPTCanvasTextPasteSourcesFromDataTransfer,
  PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL,
  type PPTCanvasRichTextPasteSource,
  type PPTCanvasTextPasteSource,
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
  hyperlinkUrl?: string
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
  hyperlinkUrl?: string
  importerId?: string
  linkRunCount: number
  name?: string
  numberedParagraphCount: number
  text: string
  textBody: PPTTextBody
  underlineRunCount: number
}
export type PPTTextPasteSourceCandidate =
  | {
      kind: 'rich-text-source'
      source: PPTRichTextPasteSource
    }
  | {
      kind: 'text-source'
      text: string
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
const PPT_RICH_TEXT_HYPERLINK_ALLOWED_SCHEMES = new Set([
  'http',
  'https',
  'mailto',
])
const PPT_RICH_TEXT_HYPERLINK_MAX_LENGTH = 2048
export const PPT_TEXT_PASTE_IMPORT_MODEL = PPT_CANVAS_TEXT_PASTE_IMPORT_MODEL

export const getPPTTextPasteSourcesFromDataTransfer =
  getPPTCanvasTextPasteSourcesFromDataTransfer

export function getPPTTextPasteSourceCandidatesFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTTextPasteSourceCandidate[] {
  return getPPTCanvasTextPasteSourceCandidatesFromDataTransfer(dataTransfer)
    .flatMap((source) => {
      const candidate = createPPTTextPasteSourceCandidate(source, {
        html: dataTransfer?.getData('text/html') ?? '',
      })

      return candidate ? [candidate] : []
    })
}

export function getPPTRichTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTRichTextPasteSource | null {
  for (const candidate of getPPTTextPasteSourceCandidatesFromDataTransfer(
    dataTransfer,
  )) {
    if (candidate.kind === 'rich-text-source') {
      return candidate.source
    }
  }

  return null
}

function createPPTTextPasteSourceCandidate(
  source: PPTCanvasTextPasteSource,
  {
    html,
  }: {
    html: string
  },
): PPTTextPasteSourceCandidate | null {
  if (source.format === 'text-plain') {
    const text = getPPTCanvasTextPasteSourceText(source)

    return text ? { kind: 'text-source', text } : null
  }

  return {
    kind: 'rich-text-source',
    source: createPPTRichTextPasteSource(source, { html }),
  }
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
    importerId: source.importerId ??
      (source.format === 'text-markdown-rich'
        ? 'canvas-rich-markdown-text'
        : 'canvas-rich-html-text'),
    italicRunCount: getPPTTextBodyRunCount(source.textBody, 'italic'),
    item: {
      ...result.item,
      ...(source.hyperlinkUrl ? { hyperlink: { url: source.hyperlinkUrl } } : {}),
      name: source.name ?? 'Rich Text',
      textBody: source.textBody,
    },
    ...(source.hyperlinkUrl ? { hyperlinkUrl: source.hyperlinkUrl } : {}),
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
  const hyperlinkUrl = getPPTRichTextSingleHyperlinkUrl(source)
  const textBody: PPTTextBody = {
    paragraphs: source.paragraphs.map((paragraph, index): PPTParagraph => {
      const bullet = listKinds[index] ?? paragraph.bullet
      const align = normalizePPTRichTextParagraphAlign(paragraph.align)

      return {
        ...(align ? { align } : {}),
        ...(bullet ? { bullet } : {}),
        ...(paragraph.lineHeight === undefined
          ? {}
          : { lineHeight: paragraph.lineHeight }),
        ...(paragraph.spacingAfter === undefined
          ? {}
          : { spacingAfter: paragraph.spacingAfter }),
        ...(paragraph.spacingBefore === undefined
          ? {}
          : { spacingBefore: paragraph.spacingBefore }),
        runs: paragraph.runs.map((run): PPTRun => ({
          ...(run.bold === undefined && !paragraph.headingLevel
            ? {}
            : { bold: run.bold ?? true }),
          ...(run.color || run.link ? { color: run.color ?? '#2563eb' } : {}),
          ...(run.fontSize === undefined ? {} : { size: run.fontSize }),
          ...(run.italic === undefined ? {} : { italic: run.italic }),
          ...(run.underline === undefined ? {} : { underline: run.underline }),
          text: run.text,
        })),
      }
    }),
  }
  const importerId = source.format === 'text-markdown-rich'
    ? 'canvas-rich-markdown-text'
    : 'canvas-rich-html-text'

  return {
    boldRunCount: getPPTTextBodyRunCount(textBody, 'bold'),
    bulletParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'bullet').length,
    format: source.format,
    ...(hyperlinkUrl ? { hyperlinkUrl } : {}),
    importerId,
    linkRunCount: source.paragraphs.reduce((count, paragraph) =>
      count + paragraph.runs.filter((run) => Boolean(run.link)).length, 0),
    name: source.format === 'text-markdown-rich'
      ? 'Markdown Text'
      : 'Rich Text',
    numberedParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'numbered').length,
    text: source.text,
    textBody,
    underlineRunCount: getPPTTextBodyRunCount(textBody, 'underline'),
  }
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

function getPPTRichTextSingleHyperlinkUrl(
  source: PPTCanvasRichTextPasteSource,
) {
  const textRuns = source.paragraphs.flatMap((paragraph) => paragraph.runs)
    .filter((run) => run.text.trim())

  if (textRuns.length === 0) {
    return null
  }

  const normalizedUrls = new Set<string>()

  for (const run of textRuns) {
    const url = normalizePPTRichTextHyperlinkUrl(run.link)

    if (!url) {
      return null
    }

    normalizedUrls.add(url)
  }

  return normalizedUrls.size === 1 ? [...normalizedUrls][0] : null
}

function normalizePPTRichTextParagraphAlign(
  align: PPTCanvasRichTextPasteSource['paragraphs'][number]['align'],
) {
  return align === 'left' || align === 'center' || align === 'right'
    ? align
    : undefined
}

function normalizePPTRichTextHyperlinkUrl(value: string | null | undefined) {
  const url = value?.trim()

  if (!url || url.length > PPT_RICH_TEXT_HYPERLINK_MAX_LENGTH) {
    return null
  }

  if (hasPPTRichTextHyperlinkControlCharacter(url)) {
    return null
  }

  const scheme = url.match(/^([a-z][a-z0-9+.-]*):/i)?.[1]?.toLowerCase()

  if (!scheme || !PPT_RICH_TEXT_HYPERLINK_ALLOWED_SCHEMES.has(scheme)) {
    return null
  }

  return url
}

function hasPPTRichTextHyperlinkControlCharacter(value: string) {
  return Array.from(value).some((character) => {
    const code = character.charCodeAt(0)

    return code <= 31 || code === 127
  })
}

function getPPTTextBodyRunCount(
  body: PPTTextBody,
  field: keyof Pick<PPTRun, 'bold' | 'color' | 'italic' | 'underline'>,
) {
  return body.paragraphs.reduce((count, paragraph) =>
    count + paragraph.runs.filter((run) => Boolean(run[field])).length, 0)
}
