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

export type PPTImageImportFormat =
  | 'file'
  | 'svg-html-img'
  | 'svg-html-inline'
  | 'svg-mime'
  | 'svg-plain'
export type PPTImageImportSource = CanvasImageImportSource & {
  format?: PPTImageImportFormat
}

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

export async function readPPTImageFileSource(file: Blob & { name?: string }) {
  const source = await readCanvasImageFileSource(file)

  return source
    ? {
        ...source,
        format: 'file' as const,
      }
    : null
}

export const getPPTImageFileFromList = getCanvasImageFileFromList
export const getPPTImageFileFromDataTransfer = getCanvasImageFileFromDataTransfer

export function getPPTSVGImageSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTImageImportSource | null {
  if (!dataTransfer) {
    return null
  }

  const svgMimeSource = getPPTSVGImageSourceFromMarkup(
    dataTransfer.getData('image/svg+xml'),
    'svg-mime',
  )

  if (svgMimeSource) {
    return svgMimeSource
  }

  const htmlSource = getPPTSVGImageSourceFromHTML(dataTransfer.getData('text/html'))

  if (htmlSource) {
    return htmlSource
  }

  return getPPTSVGImageSourceFromMarkup(
    dataTransfer.getData('text/plain'),
    'svg-plain',
  )
}

function getPPTSVGImageSourceFromHTML(value: string) {
  if (!value || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(value, 'text/html')
  const svgDataImage = doc.querySelector<HTMLImageElement>('img[src^="data:image/svg+xml"]')

  if (svgDataImage?.src) {
    const decoded = decodePPTSVGDataUrl(svgDataImage.src)

    if (decoded) {
      return getPPTSVGImageSourceFromMarkup(
        decoded,
        'svg-html-img',
        svgDataImage.alt || svgDataImage.title || undefined,
      )
    }

    return {
      dataUrl: svgDataImage.src,
      format: 'svg-html-img' as const,
      mimeType: 'image/svg+xml',
      name: getPPTSVGImportName(svgDataImage.alt || svgDataImage.title),
    }
  }

  const inlineSvg = doc.querySelector('svg')

  if (!inlineSvg) {
    return null
  }

  return getPPTSVGImageSourceFromMarkup(
    new XMLSerializer().serializeToString(inlineSvg),
    'svg-html-inline',
  )
}

function getPPTSVGImageSourceFromMarkup(
  value: string,
  format: PPTImageImportFormat,
  name?: string,
): PPTImageImportSource | null {
  const normalized = normalizePPTSVGMarkup(decodePPTSVGDataUrl(value.trim()) ?? value)

  if (!normalized) {
    return null
  }

  return {
    dataUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(normalized.markup)}`,
    format,
    mimeType: 'image/svg+xml',
    name: getPPTSVGImportName(name),
    naturalHeight: normalized.naturalHeight,
    naturalWidth: normalized.naturalWidth,
  }
}

function normalizePPTSVGMarkup(value: string) {
  if (!value.trim() || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(value.trim(), 'image/svg+xml')
  const svg = doc.documentElement

  if (
    svg.tagName.toLowerCase() !== 'svg' ||
    doc.querySelector('parsererror')
  ) {
    return null
  }

  svg.querySelectorAll('script, foreignObject').forEach((node) => node.remove())
  removeUnsafePPTSVGAttributes(svg)
  svg.querySelectorAll('*').forEach(removeUnsafePPTSVGAttributes)

  if (!svg.getAttribute('xmlns')) {
    svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg')
  }

  const size = getPPTSVGNaturalSize(svg)

  return {
    markup: new XMLSerializer().serializeToString(svg),
    naturalHeight: size?.h,
    naturalWidth: size?.w,
  }
}

function removeUnsafePPTSVGAttributes(element: Element) {
  Array.from(element.attributes).forEach((attribute) => {
    const name = attribute.name.toLowerCase()
    const value = attribute.value.trim().toLowerCase()

    if (
      name.startsWith('on') ||
      ((name === 'href' || name.endsWith(':href')) && value.startsWith('javascript:'))
    ) {
      element.removeAttribute(attribute.name)
    }
  })
}

function getPPTSVGNaturalSize(svg: Element) {
  const width = getPPTSVGLength(svg.getAttribute('width'))
  const height = getPPTSVGLength(svg.getAttribute('height'))

  if (width && height) {
    return {
      h: height,
      w: width,
    }
  }

  const viewBox = svg.getAttribute('viewBox')
    ?.trim()
    .split(/[\s,]+/)
    .map((item) => Number.parseFloat(item))

  if (
    viewBox?.length === 4 &&
    Number.isFinite(viewBox[2]) &&
    Number.isFinite(viewBox[3]) &&
    viewBox[2] > 0 &&
    viewBox[3] > 0
  ) {
    return {
      h: viewBox[3],
      w: viewBox[2],
    }
  }

  return null
}

function getPPTSVGLength(value: string | null) {
  if (!value) {
    return null
  }

  const parsed = Number.parseFloat(value)

  return Number.isFinite(parsed) && parsed > 0 ? parsed : null
}

function decodePPTSVGDataUrl(value: string) {
  const match = value.match(/^data:image\/svg\+xml[^,]*,(.*)$/i)

  if (!match) {
    return null
  }

  try {
    return value.includes(';base64,')
      ? atob(match[1])
      : decodeURIComponent(match[1])
  } catch {
    return null
  }
}

function getPPTSVGImportName(value?: string) {
  const name = value?.trim()

  return name ? `${name.replace(/\.[^.]+$/, '')}.svg` : 'clipboard.svg'
}
