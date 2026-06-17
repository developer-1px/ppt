import {
  clampPPTCanvasBoundsToFrame,
  type Point,
} from '../pptCanvasCoreAdapter'
import {
  getPPTCanvasDataImageSourceFromDataTransfer,
  getPPTCanvasImageFileFromDataTransfer,
  getPPTCanvasImageFileFromList,
  getPPTCanvasImportedImageSize,
  getPPTCanvasSVGImageSourceFromDataTransfer,
  PPT_CANVAS_IMAGE_IMPORT_MODEL,
  readPPTCanvasClipboardImageSource,
  readPPTCanvasImageFileSource,
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

export async function readPPTClipboardImageSource() {
  const source = await readPPTCanvasClipboardImageSource()

  if (!source) {
    return null
  }

  return {
    ...source,
    format: source.format ?? ('file' as const),
  }
}

export const getPPTImageFileFromList = getPPTCanvasImageFileFromList

export const getPPTImageFileFromDataTransfer =
  getPPTCanvasImageFileFromDataTransfer

export const getPPTDataImageSourceFromDataTransfer =
  getPPTCanvasDataImageSourceFromDataTransfer

export const getPPTSVGImageSourceFromDataTransfer =
  getPPTCanvasSVGImageSourceFromDataTransfer

export const resolvePPTImageSourceNaturalSize =
  resolvePPTCanvasImageSourceNaturalSize
