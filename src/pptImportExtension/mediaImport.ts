import {
  clampPPTCanvasBoundsToFrame,
  type Point,
  type Viewport,
} from '../pptCanvasCoreAdapter'
import {
  CANVAS_MEDIA_IMPORT_MODEL,
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromText,
} from 'canvas/app/media-import'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTShape,
} from '../pptModel'

export type PPTMediaImportSource = {
  title?: string
  url: string
}

export type PPTMediaImportResult = {
  importerId: string
  item: PPTShape
  source: PPTMediaImportSource
}

type CanvasMediaLinkItem = {
  fill: string
  h: number
  id: string
  shape?: 'rect'
  stroke: string
  strokeWidth?: number
  text?: string
  type: 'rect'
  w: number
  x: number
  y: number
}

const PPT_MEDIA_CARD_WIDTH = 440
const PPT_MEDIA_CARD_HEIGHT = 132

export const PPT_MEDIA_IMPORT_MODEL = CANVAS_MEDIA_IMPORT_MODEL

export const getPPTMediaSourceFromDataTransfer =
  getCanvasMediaSourceFromDataTransfer

export function createPPTMediaElement({
  createId,
  position,
  source,
  viewport,
}: {
  createId: (prefix: string) => string
  position: Point
  source: PPTMediaImportSource
  viewport: Viewport
}): PPTMediaImportResult | null {
  void viewport

  const normalized = getCanvasMediaSourceFromText(source.url)

  if (!normalized) {
    return null
  }

  const normalizedSource = {
    ...source,
    url: normalized.url,
  }
  const canvasItem = PPT_LINK_CARD_MEDIA_IMPORTER
    .createItems({
      createId,
      position,
      source: normalizedSource,
    })
    .find(isCanvasMediaLinkItem)

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
    importerId: PPT_LINK_CARD_MEDIA_IMPORTER.id,
    item: {
      cornerRadius: 16,
      fill: { color: canvasItem.fill },
      geometry,
      hyperlink: { url: normalizedSource.url },
      id: canvasItem.id,
      kind: 'shape',
      name: 'Link card',
      shape: 'rect',
      stroke: {
        color: canvasItem.stroke,
        width: canvasItem.strokeWidth ?? 2,
      },
      style: {
        color: '#1e3a8a',
        fontSize: 24,
        fontWeight: 'semibold',
        textInset: {
          bottom: 18,
          left: 22,
          right: 22,
          top: 18,
        },
      },
      textBody: createPPTTextBody(canvasItem.text ?? normalizedSource.url),
    },
    source: normalizedSource,
  }
}

const PPT_LINK_CARD_MEDIA_IMPORTER = {
  id: 'ppt-link-card',
  createItems: ({ createId, position, source }: {
    createId: (prefix: string) => string
    position: Point
    source: PPTMediaImportSource
  }) => [{
    fill: '#eff6ff',
    h: PPT_MEDIA_CARD_HEIGHT,
    id: createId('media'),
    shape: 'rect',
    stroke: '#2563eb',
    strokeWidth: 2,
    text: getPPTMediaCardText(source),
    type: 'rect',
    w: PPT_MEDIA_CARD_WIDTH,
    x: position.x - PPT_MEDIA_CARD_WIDTH / 2,
    y: position.y - PPT_MEDIA_CARD_HEIGHT / 2,
  } as const],
}

function getPPTMediaCardText(source: PPTMediaImportSource) {
  return source.title?.trim()
    ? `${source.title.trim()}\n${source.url}`
    : source.url
}

function isCanvasMediaLinkItem(value: unknown): value is CanvasMediaLinkItem {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as CanvasMediaLinkItem).type === 'rect' &&
    typeof (value as CanvasMediaLinkItem).id === 'string' &&
    typeof (value as CanvasMediaLinkItem).fill === 'string' &&
    typeof (value as CanvasMediaLinkItem).stroke === 'string' &&
    typeof (value as CanvasMediaLinkItem).x === 'number' &&
    typeof (value as CanvasMediaLinkItem).y === 'number' &&
    typeof (value as CanvasMediaLinkItem).w === 'number' &&
    typeof (value as CanvasMediaLinkItem).h === 'number'
  )
}
