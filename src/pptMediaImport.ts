import {
  clamp,
  type Point,
  type Viewport,
} from 'canvas/core'
import {
  createCanvasMediaImportItems,
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromText,
} from 'canvas/app/media-import'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTShape,
} from './pptModel'

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

export function getPPTMediaSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTMediaImportSource | null {
  return getCanvasMediaSourceFromDataTransfer(dataTransfer)
}

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
  const normalized = getCanvasMediaSourceFromText(source.url)

  if (!normalized) {
    return null
  }

  const normalizedSource = {
    ...source,
    url: normalized.url,
  }
  const result = createCanvasMediaImportItems({
    createId,
    importers: [PPT_LINK_CARD_MEDIA_IMPORTER],
    position,
    source: normalizedSource,
    viewport,
  })
  const canvasItem = result?.items.find(isCanvasMediaLinkItem)

  if (!result || !canvasItem) {
    return null
  }

  return {
    importerId: result.importerId,
    item: {
      cornerRadius: 16,
      fill: { color: canvasItem.fill },
      geometry: {
        h: canvasItem.h,
        w: canvasItem.w,
        x: clamp(canvasItem.x, 0, PPT_SLIDE_WIDTH - canvasItem.w),
        y: clamp(canvasItem.y, 0, PPT_SLIDE_HEIGHT - canvasItem.h),
      },
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
