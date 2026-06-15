import {
  clamp,
  type Point,
} from 'canvas/core'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTImage,
} from './pptModel'

export type PPTImageImportSource = {
  dataUrl: string
  mimeType: string
  name?: string
  naturalHeight?: number
  naturalWidth?: number
}

const DEFAULT_IMAGE_WIDTH = 320
const DEFAULT_IMAGE_HEIGHT = 220
const MAX_IMAGE_WIDTH = 520
const MAX_IMAGE_HEIGHT = 360

export function createPPTImportedImageElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTImageImportSource
}): PPTImage {
  const size = getPPTImportedImageSize(source)
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

export async function readPPTImageFileSource(
  file: Blob & { name?: string },
): Promise<PPTImageImportSource | null> {
  if (!isPPTImageBlob(file)) {
    return null
  }

  const dataUrl = await readPPTBlobAsDataUrl(file)
  const mimeType = file.type || getPPTImageDataUrlMimeType(dataUrl)

  if (!isPPTImageMimeType(mimeType)) {
    return null
  }

  const naturalSize = await readPPTImageNaturalSize(dataUrl)

  return {
    dataUrl,
    mimeType,
    name: file.name,
    naturalHeight: naturalSize?.h,
    naturalWidth: naturalSize?.w,
  }
}

export function getPPTImageFileFromList(files: FileList | null) {
  return Array.from(files ?? []).find(isPPTImageBlob) ?? null
}

export function getPPTImageFileFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTImageFileFromList(dataTransfer?.files ?? null)
}

function isPPTImageBlob(blob: Blob & { name?: string }): blob is File {
  return isPPTImageMimeType(blob.type)
}

function isPPTImageMimeType(value: string) {
  return /^image\/[a-z0-9.+-]+$/i.test(value)
}

function getPPTImportedImageSize({
  naturalHeight,
  naturalWidth,
}: PPTImageImportSource) {
  if (!naturalWidth || !naturalHeight) {
    return {
      h: DEFAULT_IMAGE_HEIGHT,
      w: DEFAULT_IMAGE_WIDTH,
    }
  }

  const scale = Math.min(
    1,
    MAX_IMAGE_WIDTH / naturalWidth,
    MAX_IMAGE_HEIGHT / naturalHeight,
  )

  return {
    h: Math.max(1, Math.round(naturalHeight * scale)),
    w: Math.max(1, Math.round(naturalWidth * scale)),
  }
}

function readPPTBlobAsDataUrl(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()

    reader.addEventListener('load', () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result)
        return
      }

      reject(new Error('Expected image data URL'))
    })
    reader.addEventListener('error', () => {
      reject(reader.error ?? new Error('Could not read image'))
    })
    reader.readAsDataURL(blob)
  })
}

async function readPPTImageNaturalSize(dataUrl: string) {
  try {
    const image = new Image()
    image.src = dataUrl

    await image.decode()

    return image.naturalWidth > 0 && image.naturalHeight > 0
      ? { h: image.naturalHeight, w: image.naturalWidth }
      : null
  } catch {
    return null
  }
}

function getPPTImageDataUrlMimeType(dataUrl: string) {
  return dataUrl.match(/^data:([^;,]+)/i)?.[1] ?? ''
}
