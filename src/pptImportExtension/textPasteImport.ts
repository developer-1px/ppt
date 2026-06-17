import {
  clampPPTCanvasBoundsToFrame,
  clampPPTCanvasValue,
  type Point,
  type Viewport,
} from '../pptCanvasCoreAdapter'
import {
  CANVAS_TEXT_PASTE_IMPORT_MODEL,
  createCanvasTextPasteItems,
  getCanvasRichTextPasteSourceFromDataTransfer,
  getCanvasTextPasteSourcesFromDataTransfer,
  type CanvasRichTextPasteSource,
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

type PPTTextPasteImporter =
  Parameters<typeof createCanvasTextPasteItems>[0]['importers'][number]

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

const PPT_TEXT_PASTE_WIDTH = 460
const PPT_TEXT_PASTE_LINE_HEIGHT = 38
const PPT_TEXT_PASTE_MIN_HEIGHT = 92
const PPT_TEXT_PASTE_MAX_HEIGHT = 320
export const PPT_TEXT_PASTE_IMPORT_MODEL = CANVAS_TEXT_PASTE_IMPORT_MODEL

export const getPPTTextPasteSourcesFromDataTransfer =
  getCanvasTextPasteSourcesFromDataTransfer

export function getPPTRichTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTRichTextPasteSource | null {
  const source = getCanvasRichTextPasteSourceFromDataTransfer(dataTransfer)

  return source ? createPPTRichTextPasteSource(source) : null
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
  const result = createCanvasTextPasteItems({
    createId,
    importers: [PPT_TEXT_PASTE_IMPORTER],
    position,
    text,
    viewport,
  })

  if (!result) {
    return null
  }

  const canvasItem = result.items.find(isCanvasTextPasteTextItem)

  if (!canvasItem) {
    return null
  }

  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: canvasItem.h,
      w: canvasItem.w,
      x: canvasItem.x,
      y: canvasItem.y,
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

function createPPTRichTextPasteSource(
  source: CanvasRichTextPasteSource,
): PPTRichTextPasteSource {
  const textBody: PPTTextBody = {
    paragraphs: source.paragraphs.map((paragraph): PPTParagraph => ({
      ...(paragraph.bullet ? { bullet: paragraph.bullet } : {}),
      runs: paragraph.runs.map((run): PPTRun => ({
        ...(run.bold === undefined ? {} : { bold: run.bold }),
        ...(run.color || run.link ? { color: run.color ?? '#2563eb' } : {}),
        ...(run.italic === undefined ? {} : { italic: run.italic }),
        ...(run.underline === undefined ? {} : { underline: run.underline }),
        text: run.text,
      })),
    })),
  }

  return {
    boldRunCount: getPPTTextBodyRunCount(textBody, 'bold'),
    bulletParagraphCount: textBody.paragraphs.filter((paragraph) =>
      paragraph.bullet === 'bullet').length,
    linkRunCount: source.paragraphs.reduce((count, paragraph) =>
      count + paragraph.runs.filter((run) => Boolean(run.link)).length, 0),
    text: source.text,
    textBody,
    underlineRunCount: getPPTTextBodyRunCount(textBody, 'underline'),
  }
}

function getPPTTextBodyRunCount(
  body: PPTTextBody,
  field: keyof Pick<PPTRun, 'bold' | 'color' | 'italic' | 'underline'>,
) {
  return body.paragraphs.reduce((count, paragraph) =>
    count + paragraph.runs.filter((run) => Boolean(run[field])).length, 0)
}
