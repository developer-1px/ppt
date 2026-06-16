import {
  clamp,
  type Point,
} from 'canvas/core'
import {
  getCanvasDataImageSourceFromDataTransfer,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  getCanvasImportedImageSize,
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

export function createPPTImportedImageElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTImageImportSource
}): PPTImage {
  const size = getCanvasImportedImageSize(source)
  const name = source.name?.trim() || 'Image'

  return {
    alt: name,
    crop: {
      x: 50,
      y: 50,
    },
    geometry: {
      h: size.h,
      w: size.w,
      x: clamp(center.x - size.w / 2, 0, PPT_SLIDE_WIDTH - size.w),
      y: clamp(center.y - size.h / 2, 0, PPT_SLIDE_HEIGHT - size.h),
    },
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
