import {
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromText,
} from 'canvas/app/media-import-public'

export * from 'canvas/app/media-import-public'

export type CanvasMediaImportSource = {
  title?: string
  url: string
}

export const CANVAS_MEDIA_SOURCE_IMPORT_SUPPORTED_FORMATS = [
  'text/uri-list',
  'text/plain',
  'application/json',
] as const
export const CANVAS_MEDIA_SOURCE_JSON_MIME_TYPE =
  'application/vnd.interactive-os.canvas.media-source+json'
export const CANVAS_MEDIA_SOURCE_JSON_TYPES = [
  'canvas-media-source',
  'media-source',
] as const

export type CanvasMediaObjectHyperlinkTarget<TSelection = readonly string[]> = {
  id: string
  selection: TSelection
}

export type CanvasMediaObjectHyperlinkRoute<
  TSource = CanvasMediaImportSource,
> =
  | {
    intent: {
      target: CanvasMediaObjectHyperlinkTarget
      url: string
    }
    kind: 'object-hyperlink'
  }
  | {
    kind: 'media-insert'
    source: TSource
  }

export function getCanvasMediaSourceFromJSONDataTransfer(
  dataTransfer: { getData?: (format: string) => string } | null,
): CanvasMediaImportSource | null {
  if (!dataTransfer) {
    return null
  }

  const jsonText =
    dataTransfer.getData?.(CANVAS_MEDIA_SOURCE_JSON_MIME_TYPE) ||
    dataTransfer.getData?.('application/json') ||
    ''

  if (!jsonText) {
    return null
  }

  try {
    const value = JSON.parse(jsonText) as unknown

    return getCanvasMediaSourceFromJSONValue(value)
  } catch {
    return null
  }
}

export function routeCanvasMediaSourceObjectHyperlink<
  TSource = CanvasMediaImportSource,
>({
  getTarget,
  selection = [],
  source,
}: {
  getTarget?: (input: { selection: readonly string[] }) =>
    CanvasMediaObjectHyperlinkTarget | null
  selection?: readonly string[]
  source: TSource
  targetSelection?: unknown
}): CanvasMediaObjectHyperlinkRoute<TSource> {
  const target = getTarget?.({ selection }) ?? null
  const url = isCanvasMediaRecord(source) && typeof source.url === 'string'
    ? source.url
    : ''

  return target && url
    ? {
        intent: {
          target,
          url,
        },
        kind: 'object-hyperlink',
      }
    : {
        kind: 'media-insert',
        source,
      }
}

function getCanvasMediaSourceFromJSONValue(
  value: unknown,
): CanvasMediaImportSource | null {
  if (!isCanvasMediaRecord(value)) {
    return null
  }

  const payload = isCanvasMediaRecord(value.payload) ? value.payload : value
  const url = typeof payload.url === 'string' ? payload.url : ''

  return getCanvasMediaSourceFromText(url)
}

function isCanvasMediaRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

export {
  getCanvasMediaSourceFromDataTransfer,
  getCanvasMediaSourceFromText,
}
