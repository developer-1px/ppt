import {
  clampPPTCanvasBoundsToFrame,
  type Point,
  type Viewport,
} from '../pptCanvasCoreAdapter'
import {
  getPPTCanvasMediaSourceFromDataTransfer,
  getPPTCanvasMediaSourceFromText,
  PPT_CANVAS_MEDIA_IMPORT_MODEL,
} from '../pptCanvasAppAffordanceAdapter'
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

type PPTMediaLinkItem = {
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

export const PPT_MEDIA_IMPORT_MODEL = PPT_CANVAS_MEDIA_IMPORT_MODEL

export const getPPTMediaSourceFromDataTransfer =
  getPPTCanvasMediaSourceFromDataTransfer

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

  const normalized = getPPTCanvasMediaSourceFromText(source.url)

  if (!normalized) {
    return null
  }

  const normalizedSource = {
    ...source,
    url: normalized.url,
  }
  const mediaItem = PPT_LINK_CARD_MEDIA_IMPORTER
    .createItems({
      createId,
      position,
      source: normalizedSource,
    })
    .find(isPPTMediaLinkItem)

  if (!mediaItem) {
    return null
  }

  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: mediaItem.h,
      w: mediaItem.w,
      x: mediaItem.x,
      y: mediaItem.y,
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
      fill: { color: mediaItem.fill },
      geometry,
      hyperlink: { url: normalizedSource.url },
      id: mediaItem.id,
      kind: 'shape',
      name: 'Link card',
      shape: 'rect',
      stroke: {
        color: mediaItem.stroke,
        width: mediaItem.strokeWidth ?? 2,
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
      textBody: createPPTTextBody(mediaItem.text ?? normalizedSource.url),
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

function isPPTMediaLinkItem(value: unknown): value is PPTMediaLinkItem {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as PPTMediaLinkItem).type === 'rect' &&
    typeof (value as PPTMediaLinkItem).id === 'string' &&
    typeof (value as PPTMediaLinkItem).fill === 'string' &&
    typeof (value as PPTMediaLinkItem).stroke === 'string' &&
    typeof (value as PPTMediaLinkItem).x === 'number' &&
    typeof (value as PPTMediaLinkItem).y === 'number' &&
    typeof (value as PPTMediaLinkItem).w === 'number' &&
    typeof (value as PPTMediaLinkItem).h === 'number'
  )
}
