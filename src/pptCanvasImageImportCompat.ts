import {
  getCanvasDataImageSourceFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  readCanvasImageFileSource,
  type CanvasImageImportSource,
} from 'canvas/app/image-import-public'

export * from 'canvas/app/image-import-public'

export type CanvasImagePasteReplaceTarget<TSelection = readonly string[]> = {
  id: string
  selection: TSelection
}

export type CanvasImagePasteReplaceRoute<TSource = CanvasImageImportSource> =
  | {
    intent: {
      source: TSource
      target: CanvasImagePasteReplaceTarget
    }
    kind: 'image-replace'
  }
  | {
    kind: 'image-insert'
    sources: readonly TSource[]
  }

export function getCanvasDataImageSourceFromHTML(
  value: string,
): CanvasImageImportSource | null {
  return getCanvasDataImageSourcesFromHTML(value)[0] ?? null
}

export function getCanvasHTMLDataImageSourcesFromHTML(
  value: string,
): CanvasImageImportSource[] {
  return getCanvasDataImageSourcesFromHTML(value)
}

export function getCanvasHTMLDataImageSourcesFromDataTransfer(
  dataTransfer: DataTransfer | null,
): CanvasImageImportSource[] {
  return dataTransfer
    ? getCanvasHTMLDataImageSourcesFromHTML(dataTransfer.getData('text/html'))
    : []
}

export function getCanvasImageFilesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getCanvasImageFilesFromList(dataTransfer?.files ?? null)
}

export function getCanvasImageFilesFromList(files: FileList | null) {
  return Array.from(files ?? []).filter((file) => file.type.startsWith('image/'))
}

export function getCanvasSVGImageSourceFromHTML(
  value: string,
): CanvasImageImportSource | null {
  const svg = getCanvasInlineSVGFromHTML(value)

  return svg
    ? {
        dataUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`,
        format: 'svg-html-inline',
        mimeType: 'image/svg+xml',
      }
    : null
}

export async function readCanvasImageFileSources(
  files: readonly (Blob & { name?: string })[],
): Promise<CanvasImageImportSource[]> {
  const sources = await Promise.all(files.map(readCanvasImageFileSource))

  return sources.filter((source): source is CanvasImageImportSource =>
    source !== null)
}

export function routeCanvasImagePasteReplace<TSource = CanvasImageImportSource>({
  getTarget,
  selection = [],
  source,
  sources = source === undefined ? [] : [source],
  targetSelection,
}: {
  getTarget?: (input: { selection: readonly string[] }) =>
    CanvasImagePasteReplaceTarget | null
  selection?: readonly string[]
  source?: TSource
  sources?: readonly TSource[]
  targetSelection?: unknown
}): CanvasImagePasteReplaceRoute<TSource> {
  const target = getTarget?.({ selection }) ??
    (targetSelection ? { id: '', selection: [] } : null)
  const firstSource = sources[0]

  return target && firstSource
    ? {
        intent: {
          source: firstSource,
          target,
        },
        kind: 'image-replace',
      }
    : {
        kind: 'image-insert',
        sources,
      }
}

function getCanvasDataImageSourcesFromHTML(value: string) {
  if (!value) {
    return []
  }

  if (typeof DOMParser !== 'undefined') {
    const doc = new DOMParser().parseFromString(value, 'text/html')

    return Array.from(doc.querySelectorAll('img'))
      .flatMap((image): CanvasImageImportSource[] => {
        const src = image.getAttribute('src') ?? ''

        return src.startsWith('data:') && !src.startsWith('data:image/svg+xml')
          ? [{
              dataUrl: src,
              format: 'data-url-html-img',
              mimeType: getCanvasDataUrlMimeType(src) ?? 'image/png',
              name: image.getAttribute('alt') ?? undefined,
            }]
          : []
      })
  }

  const matches = value.matchAll(/<img\b[^>]*\bsrc=["']([^"']+)["'][^>]*>/gi)

  return Array.from(matches).flatMap((match): CanvasImageImportSource[] => {
    const src = match[1] ?? ''

    return src.startsWith('data:') && !src.startsWith('data:image/svg+xml')
      ? [{
          dataUrl: src,
          format: 'data-url-html-img',
          mimeType: getCanvasDataUrlMimeType(src) ?? 'image/png',
        }]
      : []
  })
}

function getCanvasInlineSVGFromHTML(value: string) {
  if (!value) {
    return null
  }

  if (typeof DOMParser !== 'undefined') {
    const doc = new DOMParser().parseFromString(value, 'text/html')
    const svg = doc.querySelector('svg')

    return svg?.outerHTML ?? null
  }

  return value.match(/<svg\b[\s\S]*<\/svg>/i)?.[0] ?? null
}

function getCanvasDataUrlMimeType(value: string) {
  return value.match(/^data:([^;,]+)/i)?.[1] ?? null
}

export {
  getCanvasDataImageSourceFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
}
