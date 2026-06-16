import {
  clamp,
  type Point,
  type Viewport,
} from 'canvas/core'
import {
  createCanvasTextPasteItems,
  getCanvasTextPasteSourcesFromDataTransfer,
} from 'canvas/app/text-paste-import'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTTextBox,
} from './pptModel'

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
  importerId: string
  item: PPTTextBox
}

const PPT_TEXT_PASTE_WIDTH = 460
const PPT_TEXT_PASTE_LINE_HEIGHT = 38
const PPT_TEXT_PASTE_MIN_HEIGHT = 92
const PPT_TEXT_PASTE_MAX_HEIGHT = 320

export function getPPTTextPasteSourcesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getCanvasTextPasteSourcesFromDataTransfer(dataTransfer)
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
  const canvasItem = result?.items.find(isCanvasTextPasteTextItem)

  if (!result || !canvasItem) {
    return null
  }

  return {
    importerId: result.importerId,
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

const PPT_TEXT_PASTE_IMPORTER = {
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
