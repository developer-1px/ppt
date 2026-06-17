import {
  clampPPTCanvasBoundsToFrame,
  type Point,
} from '../pptCanvasCoreAdapter'
import {
  getPPTCanvasImportedImageSize,
} from '../pptCanvasAppAffordanceAdapter'
import {
  CANVAS_IMAGE_IMPORT_MODEL,
  getCanvasDataImageSourceFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  getCanvasSVGImageSourceFromDataTransfer,
  readCanvasImageFileSource,
  resolveCanvasImageSourceNaturalSize,
  type CanvasImageImportFormat,
  type CanvasImageImportSource,
} from 'canvas/app/image-import'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTImage,
} from '../pptModel'

export type PPTImageImportFormat = CanvasImageImportFormat
export type PPTImageImportSource = CanvasImageImportSource
export const PPT_IMAGE_IMPORT_MODEL = CANVAS_IMAGE_IMPORT_MODEL

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
  const source = await readCanvasImageFileSource(file)

  if (!source) {
    return null
  }

  return {
    ...source,
    format: 'file' as const,
  }
}

export const getPPTImageFileFromList = getCanvasImageFileFromList

export const getPPTImageFileFromDataTransfer =
  getCanvasImageFileFromDataTransfer

export const getPPTDataImageSourceFromDataTransfer =
  getCanvasDataImageSourceFromDataTransfer

export const getPPTSVGImageSourceFromDataTransfer =
  getCanvasSVGImageSourceFromDataTransfer

export const resolvePPTImageSourceNaturalSize =
  resolveCanvasImageSourceNaturalSize
