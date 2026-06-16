import {
  clamp,
  type Point,
} from 'canvas/core'
import {
  createCanvasImportedImageItem,
  getCanvasImageFileFromDataTransfer,
  getCanvasImageFileFromList,
  readCanvasImageFileSource,
  type CanvasImageImportSource,
} from 'canvas/app/image-import'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTImage,
} from './pptModel'

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
  const canvasImage = createCanvasImportedImageItem({
    center,
    createId,
    source,
  })
  const name = canvasImage.name?.trim() || 'Image'

  return {
    alt: canvasImage.alt?.trim() || name,
    crop: {
      x: 50,
      y: 50,
    },
    geometry: {
      h: canvasImage.h,
      w: canvasImage.w,
      x: clamp(canvasImage.x, 0, PPT_SLIDE_WIDTH - canvasImage.w),
      y: clamp(canvasImage.y, 0, PPT_SLIDE_HEIGHT - canvasImage.h),
    },
    fit: 'cover',
    id: canvasImage.id,
    kind: 'image',
    name,
    src: canvasImage.src,
  }
}

export const readPPTImageFileSource = readCanvasImageFileSource
export const getPPTImageFileFromList = getCanvasImageFileFromList
export const getPPTImageFileFromDataTransfer = getCanvasImageFileFromDataTransfer
