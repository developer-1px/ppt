import {
  clampPPTCanvasBoundsToFrame,
  type Point,
} from '../pptCanvasCoreAdapter'
import {
  createPPTCanvasExternalClipboardImagePasteActionResolver,
  createPPTCanvasExternalClipboardPasteActionPlan,
  getPPTCanvasDataImageSourceFromDataTransfer,
  getPPTCanvasImageFileFromDataTransfer,
  getPPTCanvasImageFileFromList,
  getPPTCanvasImageFilesFromDataTransfer,
  getPPTCanvasImageSourceFromDataTransfer,
  getPPTCanvasImportedImageSize,
  getPPTCanvasSVGImageSourceFromDataTransfer,
  PPT_CANVAS_IMAGE_IMPORT_MODEL,
  readPPTCanvasClipboardImageSource,
  readPPTCanvasImageFileSource,
  readPPTCanvasImageFileSources,
  resolvePPTCanvasImageSourceNaturalSize,
  type PPTCanvasImageImportFormat,
  type PPTCanvasImageImportSource,
} from '../pptCanvasAppAffordanceAdapter'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTImage,
} from '../pptModel'

export type PPTImageImportFormat = PPTCanvasImageImportFormat
export type PPTImageImportSource = PPTCanvasImageImportSource
export const PPT_IMAGE_IMPORT_MODEL = PPT_CANVAS_IMAGE_IMPORT_MODEL

const PPT_IMAGE_BATCH_GAP = 24
const PPT_IMAGE_BATCH_MAX_COLUMNS = 3

export function createPPTImportedImageElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTImageImportSource
}): PPTImage {
  const size = getPPTCanvasImportedImageSize(source)
  const name = source.name?.trim() || 'Image'
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: size.h,
      w: size.w,
      x: center.x - size.w / 2,
      y: center.y - size.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    alt: name,
    crop: {
      x: 50,
      y: 50,
    },
    geometry,
    fit: 'cover',
    id: createId('image'),
    kind: 'image',
    name,
    src: source.dataUrl,
  }
}

export function createPPTImportedImageElements({
  center,
  createId,
  sources,
}: {
  center: Point
  createId: (prefix: string) => string
  sources: readonly PPTImageImportSource[]
}): PPTImage[] {
  if (sources.length <= 1) {
    return sources.map((source) =>
      createPPTImportedImageElement({ center, createId, source })
    )
  }

  const columns = Math.min(
    PPT_IMAGE_BATCH_MAX_COLUMNS,
    Math.ceil(Math.sqrt(sources.length)),
  )
  const rows = Math.ceil(sources.length / columns)
  const sizes = sources.map((source) => getPPTCanvasImportedImageSize(source))
  const columnWidths = Array.from({ length: columns }, (_, column) =>
    Math.max(
      ...sizes
        .filter((_, index) => index % columns === column)
        .map((size) => size.w),
    )
  )
  const rowHeights = Array.from({ length: rows }, (_, row) =>
    Math.max(
      ...sizes
        .filter((_, index) => Math.floor(index / columns) === row)
        .map((size) => size.h),
    )
  )
  const totalWidth = columnWidths.reduce((sum, width) => sum + width, 0) +
    PPT_IMAGE_BATCH_GAP * Math.max(0, columns - 1)
  const totalHeight = rowHeights.reduce((sum, height) => sum + height, 0) +
    PPT_IMAGE_BATCH_GAP * Math.max(0, rows - 1)
  const origin = {
    x: center.x - totalWidth / 2,
    y: center.y - totalHeight / 2,
  }

  return sources.map((source, index) => {
    const column = index % columns
    const row = Math.floor(index / columns)
    const x = origin.x +
      columnWidths.slice(0, column).reduce((sum, width) => sum + width, 0) +
      PPT_IMAGE_BATCH_GAP * column +
      columnWidths[column] / 2
    const y = origin.y +
      rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0) +
      PPT_IMAGE_BATCH_GAP * row +
      rowHeights[row] / 2

    return createPPTImportedImageElement({
      center: { x, y },
      createId,
      source,
    })
  })
}

export async function readPPTImageFileSource(file: Blob & { name?: string }) {
  const source = await readPPTCanvasImageFileSource(file)

  if (!source) {
    return null
  }

  return {
    ...source,
    format: 'file' as const,
  }
}

export async function readPPTImageFileSources(
  files: readonly (Blob & { name?: string })[],
) {
  return (await readPPTCanvasImageFileSources(files)).map((source) => ({
    ...source,
    format: 'file' as const,
  }))
}

export async function readPPTClipboardImageSource() {
  const source = await readPPTCanvasClipboardImageSource()

  return source ? normalizePPTClipboardImageSource(source) : null
}

export async function getPPTClipboardImageImportSources() {
  return createPPTCanvasExternalClipboardPasteActionPlan<PPTImageImportSource>({
    resolvers: [
      createPPTCanvasExternalClipboardImagePasteActionResolver({
        createAction: normalizePPTClipboardImageSource,
        readImageSource: readPPTCanvasClipboardImageSource,
      }),
    ],
  })
}

function normalizePPTClipboardImageSource(
  source: PPTCanvasImageImportSource,
): PPTImageImportSource {
  return {
    ...source,
    format: source.format ?? ('file' as const),
  }
}

export const getPPTImageFileFromList = getPPTCanvasImageFileFromList

export const getPPTImageFileFromDataTransfer =
  getPPTCanvasImageFileFromDataTransfer

export const getPPTImageFilesFromDataTransfer =
  getPPTCanvasImageFilesFromDataTransfer

export const getPPTDataImageSourceFromDataTransfer =
  getPPTCanvasDataImageSourceFromDataTransfer

export const getPPTImageSourceFromDataTransfer =
  getPPTCanvasImageSourceFromDataTransfer

export const getPPTSVGImageSourceFromDataTransfer =
  getPPTCanvasSVGImageSourceFromDataTransfer

export function shouldResolvePPTImageSourceNaturalSize(
  source: PPTImageImportSource,
) {
  return source.format === 'data-url-html-img' ||
    source.format === 'data-url-plain'
}

export const resolvePPTImageSourceNaturalSize =
  resolvePPTCanvasImageSourceNaturalSize
