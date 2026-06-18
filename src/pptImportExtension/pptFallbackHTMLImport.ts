import {
  clampPPTCanvasBoundsToFrame,
  clampPPTCanvasValue,
  type Point,
} from '../pptCanvasCoreAdapter'
import {
  createPPTTextBody,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTFill,
  type PPTImage,
  type PPTImageCrop,
  type PPTImageFit,
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTShapeKind,
  type PPTStroke,
  type PPTTable,
  type PPTTextBody,
  type PPTTextBox,
  type PPTTextStyle,
} from '../pptModel'
import {
  getPPTHTMLDataImageSourcesFromHTML,
  getPPTImageSourceFromDataTransfer,
  type PPTImageImportSource,
} from './imageImport'
import {
  getPPTTableSourceFromHTML,
  normalizePPTTableRows,
  type PPTTableImportSource,
} from './tableImport'

export const PPT_FALLBACK_HTML_IMPORT_MODEL = 'ppt-fallback-html-import'

export type PPTFallbackHTMLShapeSource = {
  cornerRadius?: number
  fill: PPTFill
  geometry: {
    h: number
    w: number
    x?: number
    y?: number
  }
  name: string
  shape: PPTShapeKind
  sourceObjectId?: string
  stroke?: PPTStroke
  style: PPTTextStyle
  textBody?: PPTTextBody
}

export type PPTFallbackHTMLTextSource = {
  geometry: {
    h: number
    w: number
    x?: number
    y?: number
  }
  name: string
  sourceObjectId?: string
  style: PPTTextStyle
  textBody: PPTTextBody
}

export type PPTFallbackHTMLImageSource = {
  alt: string
  crop: PPTImageCrop
  fit: PPTImageFit
  geometry: {
    h: number
    w: number
    x?: number
    y?: number
  }
  image: PPTImageImportSource
  name: string
  sourceObjectId?: string
}

export type PPTFallbackHTMLTableSource = {
  geometry: {
    h: number
    w: number
    x?: number
    y?: number
  }
  name: string
  sourceObjectId?: string
  table: PPTTableImportSource
}

export type PPTFallbackHTMLSelectionItemSource =
  | {
      kind: PPTImage['kind']
      source: PPTFallbackHTMLImageSource
    }
  | {
      kind: PPTTable['kind']
      source: PPTFallbackHTMLTableSource
    }
  | {
      kind: PPTShape['kind']
      source: PPTFallbackHTMLShapeSource
    }
  | {
      kind: PPTTextBox['kind']
      source: PPTFallbackHTMLTextSource
    }

export type PPTFallbackHTMLSelectionSource = {
  items: PPTFallbackHTMLSelectionItemSource[]
  name: string
}

export type PPTFallbackHTMLImportEffect = {
  format: 'text-html-ppt-fallback'
  kind:
    | PPTImage['kind']
    | PPTShape['kind']
    | PPTTable['kind']
    | PPTTextBox['kind']
    | 'selection'
  model: typeof PPT_FALLBACK_HTML_IMPORT_MODEL
  name: string
  objectCount?: number
  shape?: PPTShapeKind
  sourceObjectId?: string
  sourceObjectIds?: string[]
}

const PPT_FALLBACK_SHAPE_MIN_SIZE = 48
const PPT_FALLBACK_SHAPE_MAX_WIDTH = 980
const PPT_FALLBACK_SHAPE_MAX_HEIGHT = 560
const PPT_FALLBACK_IMAGE_MIN_SIZE = 24
const PPT_FALLBACK_IMAGE_MAX_WIDTH = 980
const PPT_FALLBACK_IMAGE_MAX_HEIGHT = 560
const PPT_FALLBACK_TABLE_MIN_SIZE = 24
const PPT_FALLBACK_TABLE_MAX_WIDTH = 980
const PPT_FALLBACK_TABLE_MAX_HEIGHT = 560
const PPT_FALLBACK_TEXT_MIN_SIZE = 24
const PPT_FALLBACK_TEXT_MAX_WIDTH = 980
const PPT_FALLBACK_TEXT_MAX_HEIGHT = 560
const PPT_FALLBACK_SELECTION_GRID_GAP = 24
const PPT_FALLBACK_SELECTION_GRID_MAX_COLUMNS = 2

export function getPPTFallbackHTMLShapeSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTFallbackHTMLShapeSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
  )
}

export function getPPTFallbackHTMLTextSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTFallbackHTMLTextSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
  )
}

export function getPPTFallbackHTMLImageSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  const image = getPPTImageSourceFromDataTransfer(dataTransfer)

  return getPPTFallbackHTMLImageSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
    image,
  )
}

export function getPPTFallbackHTMLTableSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTFallbackHTMLTableSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
  )
}

export function getPPTFallbackHTMLSelectionSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  const image = getPPTImageSourceFromDataTransfer(dataTransfer)

  return getPPTFallbackHTMLSelectionSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
    image,
  )
}

export function getPPTFallbackHTMLSelectionSourceFromHTML(
  html: string,
  image: PPTImageImportSource | null = null,
): PPTFallbackHTMLSelectionSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const items = getPPTFallbackHTMLSelectionItemSources(doc, image)

  return items.length > 1
    ? {
        items,
        name: 'PPT HTML Selection',
      }
    : null
}

export function getPPTFallbackHTMLShapeSourceFromHTML(
  html: string,
): PPTFallbackHTMLShapeSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const element = doc.querySelector<HTMLElement>('[data-ppt-selection-shape]')

  if (!element) {
    return null
  }

  return getPPTFallbackHTMLShapeSourceFromElement(element)
}

export function getPPTFallbackHTMLTextSourceFromHTML(
  html: string,
): PPTFallbackHTMLTextSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const element = [...doc.querySelectorAll<HTMLElement>(
    '[data-ppt-selection-text-body]',
  )].find(isPPTFallbackHTMLStandaloneTextElement)

  return element ? getPPTFallbackHTMLTextSourceFromElement(element) : null
}

export function getPPTFallbackHTMLImageSourceFromHTML(
  html: string,
  image: PPTImageImportSource | null,
): PPTFallbackHTMLImageSource | null {
  if (!html || !image || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const element = doc.querySelector<HTMLElement>('[data-ppt-selection-image]')

  return element ? getPPTFallbackHTMLImageSourceFromElement(element, image) : null
}

export function getPPTFallbackHTMLTableSourceFromHTML(
  html: string,
): PPTFallbackHTMLTableSource | null {
  if (!html || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(html, 'text/html')
  const element = doc.querySelector<HTMLElement>('[data-ppt-selection-table]')

  return element ? getPPTFallbackHTMLTableSourceFromElement(element) : null
}

function getPPTFallbackHTMLSelectionItemSources(
  doc: Document,
  image: PPTImageImportSource | null = null,
): PPTFallbackHTMLSelectionItemSource[] {
  const markedItems = getPPTFallbackHTMLMarkedSelectionItemSources(doc, image)

  return markedItems.length > 0
    ? markedItems
    : getPPTFallbackHTMLExternalSelectionItemSources(doc)
}

function getPPTFallbackHTMLMarkedSelectionItemSources(
  doc: Document,
  image: PPTImageImportSource | null = null,
): PPTFallbackHTMLSelectionItemSource[] {
  const imageElementCount = doc.querySelectorAll(
    '[data-ppt-selection-image]',
  ).length

  return [...doc.querySelectorAll<HTMLElement>(
    '[data-ppt-selection-image], [data-ppt-selection-shape], [data-ppt-selection-table], [data-ppt-selection-text-body]',
  )]
    .map((element): PPTFallbackHTMLSelectionItemSource | null => {
      if (element.hasAttribute('data-ppt-selection-image')) {
        const source = getPPTFallbackHTMLImageSourceFromElement(
          element,
          imageElementCount === 1 ? image : null,
        )

        return source ? { kind: 'image', source } : null
      }

      if (element.hasAttribute('data-ppt-selection-shape')) {
        const source = getPPTFallbackHTMLShapeSourceFromElement(element)

        return source ? { kind: 'shape', source } : null
      }

      if (element.hasAttribute('data-ppt-selection-table')) {
        const source = getPPTFallbackHTMLTableSourceFromElement(element)

        return source ? { kind: 'table', source } : null
      }

      if (isPPTFallbackHTMLStandaloneTextElement(element)) {
        const source = getPPTFallbackHTMLTextSourceFromElement(element)

        return source ? { kind: 'textBox', source } : null
      }

      return null
    })
    .filter((item): item is PPTFallbackHTMLSelectionItemSource =>
      item !== null)
}

function getPPTFallbackHTMLExternalSelectionItemSources(
  doc: Document,
): PPTFallbackHTMLSelectionItemSource[] {
  const seenImageDataUrls = new Set<string>()

  return [...doc.querySelectorAll<HTMLElement>('img[src^="data:image/"], table')]
    .map((element): PPTFallbackHTMLSelectionItemSource | null => {
      if (element instanceof HTMLImageElement) {
        const source = getPPTFallbackHTMLExternalImageSourceFromElement(
          element,
          seenImageDataUrls,
        )

        return source ? { kind: 'image', source } : null
      }

      if (element instanceof HTMLTableElement) {
        const source = getPPTFallbackHTMLTableSourceFromElement(element)

        return source ? { kind: 'table', source } : null
      }

      return null
    })
    .filter((item): item is PPTFallbackHTMLSelectionItemSource =>
      item !== null)
}

function isPPTFallbackHTMLStandaloneTextElement(element: HTMLElement) {
  return !element.closest('[data-ppt-selection-shape]') &&
    element.getAttribute('data-ppt-selection-shape') === null
}

function getPPTFallbackHTMLExternalImageSourceFromElement(
  element: HTMLImageElement,
  seenDataUrls?: Set<string>,
) {
  const image = getPPTFallbackHTMLImageImportSourceFromHTML(element.outerHTML)

  if (image && seenDataUrls?.has(image.dataUrl)) {
    return null
  }

  if (image) {
    seenDataUrls?.add(image.dataUrl)
  }

  return getPPTFallbackHTMLImageSourceFromElement(element, image)
}

function getPPTFallbackHTMLImageSourceFromElement(
  element: HTMLElement,
  image?: PPTImageImportSource | null,
): PPTFallbackHTMLImageSource | null {
  const img = element.matches('img')
    ? element as HTMLImageElement
    : element.querySelector<HTMLImageElement>('img')
  const imageSource = image ??
    (img ? getPPTFallbackHTMLImageImportSourceFromHTML(img.outerHTML) : null)

  if (!img || !imageSource) {
    return null
  }

  const rawStyle = element.getAttribute('style') ?? ''
  const imageStyle = img.getAttribute('style') ?? ''
  const sourceObjectId =
    element.getAttribute('data-ppt-selection-object')?.trim() || undefined
  const alt = img.getAttribute('alt')?.trim() ||
    element.querySelector('figcaption')?.textContent?.trim() ||
    imageSource.name?.trim() ||
    'Image'

  return {
    alt,
    crop: {
      x: parsePPTFallbackHTMLImageCropValue(
        element,
        'data-ppt-selection-image-crop-x',
        imageStyle,
        'x',
      ),
      y: parsePPTFallbackHTMLImageCropValue(
        element,
        'data-ppt-selection-image-crop-y',
        imageStyle,
        'y',
      ),
    },
    fit: parsePPTFallbackHTMLImageFit(
      element.getAttribute('data-ppt-selection-image-fit') ??
        getPPTFallbackHTMLStyleValue(imageStyle, 'object-fit'),
    ),
    geometry: {
      h: clampPPTCanvasValue(
        parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-h') ??
          parsePPTFallbackHTMLPixelStyle(rawStyle, 'height') ??
          imageSource.naturalHeight ??
          220,
        PPT_FALLBACK_IMAGE_MIN_SIZE,
        PPT_FALLBACK_IMAGE_MAX_HEIGHT,
      ),
      ...parsePPTFallbackHTMLPositionAttributes(element),
      w: clampPPTCanvasValue(
        parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-w') ??
          parsePPTFallbackHTMLPixelStyle(rawStyle, 'width') ??
          imageSource.naturalWidth ??
          320,
        PPT_FALLBACK_IMAGE_MIN_SIZE,
        PPT_FALLBACK_IMAGE_MAX_WIDTH,
      ),
    },
    image: {
      ...imageSource,
      name: alt,
    },
    name: alt,
    ...(sourceObjectId ? { sourceObjectId } : {}),
  }
}

function getPPTFallbackHTMLTableSourceFromElement(
  element: HTMLElement,
): PPTFallbackHTMLTableSource | null {
  const table = getPPTTableSourceFromHTML(element.outerHTML)

  if (!table) {
    return null
  }

  const rawStyle = element.getAttribute('style') ?? ''
  const sourceObjectId =
    element.getAttribute('data-ppt-selection-object')?.trim() ||
    element.getAttribute('data-ppt-table-export')?.trim() ||
    undefined

  return {
    geometry: {
      h: clampPPTCanvasValue(
        parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-h') ??
          parsePPTFallbackHTMLPixelStyle(rawStyle, 'height') ??
          180,
        PPT_FALLBACK_TABLE_MIN_SIZE,
        PPT_FALLBACK_TABLE_MAX_HEIGHT,
      ),
      ...parsePPTFallbackHTMLPositionAttributes(element),
      w: clampPPTCanvasValue(
        parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-w') ??
          parsePPTFallbackHTMLPixelStyle(rawStyle, 'width') ??
          420,
        PPT_FALLBACK_TABLE_MIN_SIZE,
        PPT_FALLBACK_TABLE_MAX_WIDTH,
      ),
    },
    name: table.name ?? 'PPT HTML Table',
    ...(sourceObjectId ? { sourceObjectId } : {}),
    table,
  }
}

function getPPTFallbackHTMLShapeSourceFromElement(
  element: HTMLElement,
): PPTFallbackHTMLShapeSource | null {
  const shape = parsePPTFallbackShapeKind(
    element.getAttribute('data-ppt-selection-shape'),
  )

  if (!shape) {
    return null
  }

  const rawStyle = element.getAttribute('style') ?? ''
  const textBody = getPPTFallbackHTMLTextBody(element)
  const sourceObjectId =
    element.getAttribute('data-ppt-selection-object')?.trim() || undefined

  return {
    ...(shape === 'rect'
      ? { cornerRadius: parsePPTFallbackHTMLPixelStyle(rawStyle, 'border-radius') }
      : {}),
    fill: {
      color: parsePPTFallbackHTMLColor(
        getPPTFallbackHTMLStyleValue(rawStyle, 'background'),
        '#e0f2fe',
      ),
    },
    geometry: {
      h: clampPPTCanvasValue(
        parsePPTFallbackHTMLPixelStyle(rawStyle, 'height') ?? 160,
        PPT_FALLBACK_SHAPE_MIN_SIZE,
        PPT_FALLBACK_SHAPE_MAX_HEIGHT,
      ),
      ...parsePPTFallbackHTMLPositionAttributes(element),
      w: clampPPTCanvasValue(
        parsePPTFallbackHTMLPixelStyle(rawStyle, 'width') ?? 320,
        PPT_FALLBACK_SHAPE_MIN_SIZE,
        PPT_FALLBACK_SHAPE_MAX_WIDTH,
      ),
    },
    name: 'PPT HTML Shape',
    shape,
    ...(sourceObjectId ? { sourceObjectId } : {}),
    stroke: parsePPTFallbackHTMLBorder(rawStyle),
    style: {
      color: parsePPTFallbackHTMLColor(
        getPPTFallbackHTMLStyleValue(rawStyle, 'color'),
        '#0f172a',
      ),
      fontFamily: normalizePPTFallbackHTMLFontFamily(
        getPPTFallbackHTMLStyleValue(rawStyle, 'font-family'),
      ),
      fontSize: parsePPTFallbackHTMLPixelStyle(rawStyle, 'font-size') ?? 24,
      fontWeight: parsePPTFallbackHTMLFontWeight(
        getPPTFallbackHTMLStyleValue(rawStyle, 'font-weight'),
      ),
      textInset: parsePPTFallbackHTMLPadding(rawStyle, 18),
      verticalAlign: parsePPTFallbackHTMLVerticalAlign(
        getPPTFallbackHTMLStyleValue(rawStyle, 'align-items'),
      ),
    },
    ...(textBody ? { textBody } : {}),
  }
}

function getPPTFallbackHTMLTextSourceFromElement(
  element: HTMLElement,
): PPTFallbackHTMLTextSource | null {
  const textBody = getPPTFallbackHTMLTextBody(element)

  if (!textBody) {
    return null
  }

  const rawStyle = element.getAttribute('style') ?? ''
  const sourceObjectId =
    element.getAttribute('data-ppt-selection-object')?.trim() || undefined

  return {
    geometry: {
      h: clampPPTCanvasValue(
        parsePPTFallbackHTMLPixelStyle(rawStyle, 'height') ??
          getPPTFallbackHTMLTextHeight(textBody),
        PPT_FALLBACK_TEXT_MIN_SIZE,
        PPT_FALLBACK_TEXT_MAX_HEIGHT,
      ),
      ...parsePPTFallbackHTMLPositionAttributes(element),
      w: clampPPTCanvasValue(
        parsePPTFallbackHTMLPixelStyle(rawStyle, 'width') ?? 460,
        PPT_FALLBACK_TEXT_MIN_SIZE,
        PPT_FALLBACK_TEXT_MAX_WIDTH,
      ),
    },
    name: 'PPT HTML Text',
    ...(sourceObjectId ? { sourceObjectId } : {}),
    style: {
      color: parsePPTFallbackHTMLColor(
        getPPTFallbackHTMLStyleValue(rawStyle, 'color'),
        '#111827',
      ),
      fontFamily: normalizePPTFallbackHTMLFontFamily(
        getPPTFallbackHTMLStyleValue(rawStyle, 'font-family'),
      ),
      fontSize: parsePPTFallbackHTMLPixelStyle(rawStyle, 'font-size') ?? 24,
      fontWeight: parsePPTFallbackHTMLFontWeight(
        getPPTFallbackHTMLStyleValue(rawStyle, 'font-weight'),
      ),
      textInset: parsePPTFallbackHTMLPadding(rawStyle, 0),
      verticalAlign: parsePPTFallbackHTMLVerticalAlign(
        getPPTFallbackHTMLStyleValue(rawStyle, 'align-items'),
      ),
    },
    textBody,
  }
}

export function createPPTFallbackHTMLSelectionElements({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTFallbackHTMLSelectionSource
}) {
  const centers = getPPTFallbackHTMLSelectionItemCenters(source, center)

  return source.items.map((item, index) => {
    if (item.kind === 'image') {
      return createPPTFallbackHTMLImageElement({
        center: centers[index],
        createId,
        source: item.source,
      })
    }

    if (item.kind === 'shape') {
      return createPPTFallbackHTMLShapeElement({
        center: centers[index],
        createId,
        source: item.source,
      })
    }

    if (item.kind === 'table') {
      return createPPTFallbackHTMLTableElement({
        center: centers[index],
        createId,
        source: item.source,
      })
    }

    return createPPTFallbackHTMLTextElement({
      center: centers[index],
      createId,
      source: item.source,
    })
  })
}

export function createPPTFallbackHTMLTableElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTFallbackHTMLTableSource
}): PPTTable {
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: source.geometry.h,
      w: source.geometry.w,
      x: center.x - source.geometry.w / 2,
      y: center.y - source.geometry.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    geometry,
    id: createId('table'),
    kind: 'table',
    name: source.name,
    rows: normalizePPTTableRows(source.table.rows),
  }
}

export function createPPTFallbackHTMLImageElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTFallbackHTMLImageSource
}): PPTImage {
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: source.geometry.h,
      w: source.geometry.w,
      x: center.x - source.geometry.w / 2,
      y: center.y - source.geometry.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    alt: source.alt,
    crop: source.crop,
    fit: source.fit,
    geometry,
    id: createId('image'),
    kind: 'image',
    name: source.name,
    src: source.image.dataUrl,
  }
}

export function createPPTFallbackHTMLShapeElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTFallbackHTMLShapeSource
}): PPTShape {
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: source.geometry.h,
      w: source.geometry.w,
      x: center.x - source.geometry.w / 2,
      y: center.y - source.geometry.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    ...(source.cornerRadius === undefined
      ? {}
      : { cornerRadius: source.cornerRadius }),
    fill: source.fill,
    geometry,
    id: createId('shape'),
    kind: 'shape',
    name: source.name,
    shape: source.shape,
    ...(source.stroke ? { stroke: source.stroke } : {}),
    style: source.style,
    ...(source.textBody ? { textBody: source.textBody } : {}),
    ...(source.textBody ? { textAutoFit: 'resizeShapeToFitText' as const } : {}),
  }
}

export function createPPTFallbackHTMLTextElement({
  center,
  createId,
  source,
}: {
  center: Point
  createId: (prefix: string) => string
  source: PPTFallbackHTMLTextSource
}): PPTTextBox {
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: source.geometry.h,
      w: source.geometry.w,
      x: center.x - source.geometry.w / 2,
      y: center.y - source.geometry.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    geometry,
    id: createId('text'),
    kind: 'textBox',
    name: source.name,
    style: source.style,
    textBody: source.textBody,
  }
}

export function createPPTFallbackHTMLSelectionImportEffect({
  elements,
  source,
}: {
  elements: readonly (PPTImage | PPTShape | PPTTable | PPTTextBox)[]
  source: PPTFallbackHTMLSelectionSource
}): PPTFallbackHTMLImportEffect {
  const sourceObjectIds = source.items
    .map((item) => item.source.sourceObjectId)
    .filter((id): id is string => Boolean(id))

  return {
    format: 'text-html-ppt-fallback',
    kind: 'selection',
    model: PPT_FALLBACK_HTML_IMPORT_MODEL,
    name: source.name,
    objectCount: elements.length,
    ...(sourceObjectIds.length > 0 ? { sourceObjectIds } : {}),
  }
}

export function createPPTFallbackHTMLImportEffect({
  element,
  source,
}: {
  element: PPTImage | PPTShape | PPTTable | PPTTextBox
  source:
    | PPTFallbackHTMLImageSource
    | PPTFallbackHTMLShapeSource
    | PPTFallbackHTMLTableSource
    | PPTFallbackHTMLTextSource
}): PPTFallbackHTMLImportEffect {
  return {
    format: 'text-html-ppt-fallback',
    kind: element.kind,
    model: PPT_FALLBACK_HTML_IMPORT_MODEL,
    name: element.name,
    ...(element.kind === 'shape' ? { shape: element.shape } : {}),
    ...(source.sourceObjectId ? { sourceObjectId: source.sourceObjectId } : {}),
  }
}

function getPPTFallbackHTMLImageImportSourceFromHTML(
  html: string,
): PPTImageImportSource | null {
  const [source = null] = getPPTHTMLDataImageSourcesFromHTML(html)

  if (source) {
    return source
  }

  const dataTransfer = {
    getData: (type: string) => type === 'text/html' ? html : '',
  } as DataTransfer

  return getPPTImageSourceFromDataTransfer(dataTransfer)
}

function parsePPTFallbackHTMLImageFit(value: string | null): PPTImageFit {
  return value === 'contain' ? 'contain' : 'cover'
}

function parsePPTFallbackHTMLImageCropValue(
  element: HTMLElement,
  attribute: string,
  imageStyle: string,
  axis: 'x' | 'y',
) {
  const attributeValue = parsePPTFallbackHTMLNumberAttribute(element, attribute)

  if (attributeValue !== undefined) {
    return clampPPTCanvasValue(attributeValue, 0, 100)
  }

  const position = getPPTFallbackHTMLStyleValue(imageStyle, 'object-position')
  const [xValue = '50%', yValue = '50%'] = position.split(/\s+/)
  const rawValue = axis === 'x' ? xValue : yValue
  const parsed = Number(rawValue.match(/-?\d+(?:\.\d+)?/)?.[0] ?? NaN)

  return Number.isFinite(parsed)
    ? clampPPTCanvasValue(parsed, 0, 100)
    : 50
}

function getPPTFallbackHTMLSelectionItemCenters(
  source: PPTFallbackHTMLSelectionSource,
  center: Point,
) {
  if (source.items.every((item) =>
    item.source.geometry.x !== undefined &&
    item.source.geometry.y !== undefined)) {
    const bounds = source.items.reduce(
      (current, item) => {
        const x = item.source.geometry.x ?? 0
        const y = item.source.geometry.y ?? 0

        return {
          maxX: Math.max(current.maxX, x + item.source.geometry.w),
          maxY: Math.max(current.maxY, y + item.source.geometry.h),
          minX: Math.min(current.minX, x),
          minY: Math.min(current.minY, y),
        }
      },
      {
        maxX: Number.NEGATIVE_INFINITY,
        maxY: Number.NEGATIVE_INFINITY,
        minX: Number.POSITIVE_INFINITY,
        minY: Number.POSITIVE_INFINITY,
      },
    )
    const offset = {
      x: center.x - (bounds.minX + (bounds.maxX - bounds.minX) / 2),
      y: center.y - (bounds.minY + (bounds.maxY - bounds.minY) / 2),
    }

    return source.items.map((item) => ({
      x: (item.source.geometry.x ?? 0) + item.source.geometry.w / 2 + offset.x,
      y: (item.source.geometry.y ?? 0) + item.source.geometry.h / 2 + offset.y,
    }))
  }

  const columns = Math.min(
    PPT_FALLBACK_SELECTION_GRID_MAX_COLUMNS,
    Math.ceil(Math.sqrt(source.items.length)),
  )
  const rows = Math.ceil(source.items.length / columns)
  const columnWidths = Array.from({ length: columns }, (_, column) =>
    Math.max(
      ...source.items
        .filter((_, index) => index % columns === column)
        .map((item) => item.source.geometry.w),
    )
  )
  const rowHeights = Array.from({ length: rows }, (_, row) =>
    Math.max(
      ...source.items
        .filter((_, index) => Math.floor(index / columns) === row)
        .map((item) => item.source.geometry.h),
    )
  )
  const totalWidth = columnWidths.reduce((sum, width) => sum + width, 0) +
    PPT_FALLBACK_SELECTION_GRID_GAP * Math.max(0, columns - 1)
  const totalHeight = rowHeights.reduce((sum, height) => sum + height, 0) +
    PPT_FALLBACK_SELECTION_GRID_GAP * Math.max(0, rows - 1)
  const origin = {
    x: center.x - totalWidth / 2,
    y: center.y - totalHeight / 2,
  }

  return source.items.map((_, index) => {
    const column = index % columns
    const row = Math.floor(index / columns)

    return {
      x: origin.x +
        columnWidths.slice(0, column).reduce((sum, width) => sum + width, 0) +
        PPT_FALLBACK_SELECTION_GRID_GAP * column +
        columnWidths[column] / 2,
      y: origin.y +
        rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0) +
        PPT_FALLBACK_SELECTION_GRID_GAP * row +
        rowHeights[row] / 2,
    }
  })
}

function parsePPTFallbackShapeKind(value: string | null): PPTShapeKind | null {
  return value === 'rect' || value === 'ellipse' || value === 'diamond'
    ? value
    : null
}

function getPPTFallbackHTMLStyleValue(style: string, property: string) {
  const escapedProperty = property.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const match = style.match(new RegExp(
    `(?:^|;)\\s*${escapedProperty}\\s*:\\s*([^;]+)`,
    'i',
  ))

  return match?.[1]?.trim() ?? ''
}

function parsePPTFallbackHTMLPixelStyle(
  style: string,
  property: string,
) {
  const value = getPPTFallbackHTMLStyleValue(style, property)
  const match = value.match(/-?\d+(?:\.\d+)?/)

  return match ? Number(match[0]) : undefined
}

function parsePPTFallbackHTMLPositionAttributes(element: HTMLElement) {
  const x = parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-x')
  const y = parsePPTFallbackHTMLNumberAttribute(element, 'data-ppt-selection-y')

  return {
    ...(x === undefined ? {} : { x }),
    ...(y === undefined ? {} : { y }),
  }
}

function parsePPTFallbackHTMLNumberAttribute(
  element: HTMLElement,
  attribute: string,
) {
  const rawValue = element.getAttribute(attribute)

  if (rawValue === null || rawValue.trim() === '') {
    return undefined
  }

  const value = Number(rawValue)

  return Number.isFinite(value) ? value : undefined
}

function parsePPTFallbackHTMLBorder(style: string): PPTStroke | undefined {
  const value = getPPTFallbackHTMLStyleValue(style, 'border')

  if (!value || value === 'none') {
    return undefined
  }

  const width = Number(value.match(/(\d+(?:\.\d+)?)px/)?.[1] ?? 0)
  const color = parsePPTFallbackHTMLColor(value, '')

  if (!width || !color) {
    return undefined
  }

  return {
    color,
    dash: value.includes('dashed')
      ? 'dash'
      : value.includes('dotted')
        ? 'dot'
        : undefined,
    width,
  }
}

function parsePPTFallbackHTMLColor(value: string, fallback: string) {
  const hex = value.match(/#[\da-f]{3,8}\b/i)?.[0]

  if (hex) {
    return normalizePPTFallbackHTMLHexColor(hex)
  }

  const rgb = value.match(/rgba?\(([^)]+)\)/i)

  if (!rgb) {
    return fallback
  }

  const [r = 0, g = 0, b = 0] = rgb[1]
    .split(/[,\s/]+/)
    .map((part) => Number(part.trim()))
    .filter((part) => Number.isFinite(part))

  return `#${[r, g, b]
    .map((part) =>
      Math.max(0, Math.min(255, Math.round(part)))
        .toString(16)
        .padStart(2, '0'))
    .join('')}`
}

function normalizePPTFallbackHTMLHexColor(value: string) {
  if (value.length === 4) {
    return `#${value[1]}${value[1]}${value[2]}${value[2]}${value[3]}${value[3]}`
      .toLowerCase()
  }

  return value.slice(0, 7).toLowerCase()
}

function parsePPTFallbackHTMLFontWeight(
  value: string,
): PPTTextStyle['fontWeight'] {
  if (value === '700' || value === 'bold') {
    return 'bold'
  }

  if (value === '600' || value === 'semibold') {
    return 'semibold'
  }

  return 'regular'
}

function normalizePPTFallbackHTMLFontFamily(value: string) {
  const [firstFamily = 'Inter'] = value.split(',')

  return firstFamily.trim().replace(/^["']|["']$/g, '') || 'Inter'
}

function parsePPTFallbackHTMLPadding(style: string, fallbackValue: number) {
  const value = getPPTFallbackHTMLStyleValue(style, 'padding')
  const parts = value
    .split(/\s+/)
    .map((part) => Number(part.match(/-?\d+(?:\.\d+)?/)?.[0] ?? NaN))
    .filter((part) => Number.isFinite(part))

  if (parts.length === 0) {
    return {
      bottom: fallbackValue,
      left: fallbackValue,
      right: fallbackValue,
      top: fallbackValue,
    }
  }

  const [top, right = top, bottom = top, left = right] = parts

  return {
    bottom,
    left,
    right,
    top,
  }
}

function getPPTFallbackHTMLTextHeight(body: PPTTextBody) {
  return Math.max(1, body.paragraphs.length) * 38 + 32
}

function parsePPTFallbackHTMLVerticalAlign(
  value: string,
): PPTTextStyle['verticalAlign'] {
  if (value === 'flex-end') {
    return 'bottom'
  }

  if (value === 'center') {
    return 'middle'
  }

  return 'top'
}

function getPPTFallbackHTMLTextBody(element: HTMLElement): PPTTextBody | null {
  const root = element.querySelector<HTMLElement>('[data-ppt-selection-text-body]')
    ?? element
  const paragraphElements = [
    ...root.querySelectorAll<HTMLElement>('[data-ppt-selection-paragraph]'),
  ]

  if (paragraphElements.length === 0) {
    const text = root.textContent?.trim() ?? ''

    return text ? createPPTTextBody(text) : null
  }

  const paragraphs = paragraphElements
    .map((paragraphElement): PPTParagraph | null => {
      const runs = getPPTFallbackHTMLRuns(paragraphElement)
      const text = runs.map((run) => run.text).join('').trim()

      if (!text) {
        return null
      }

      return {
        ...getPPTFallbackHTMLParagraphListAttribute(paragraphElement),
        ...getPPTFallbackHTMLParagraphAttributes(paragraphElement),
        runs,
      }
    })
    .filter((paragraph): paragraph is PPTParagraph => paragraph !== null)

  return paragraphs.length > 0 ? { paragraphs } : null
}

function getPPTFallbackHTMLParagraphListAttribute(
  element: HTMLElement,
): Pick<PPTParagraph, 'bullet'> {
  if (element.tagName.toLowerCase() !== 'li') {
    return {}
  }

  return {
    bullet: element.parentElement?.tagName.toLowerCase() === 'ol'
      ? 'numbered'
      : 'bullet',
  }
}

function getPPTFallbackHTMLRuns(root: HTMLElement): PPTRun[] {
  const runs: PPTRun[] = []

  collectPPTFallbackHTMLRuns(root, {}, runs)

  return runs.length > 0 ? runs : [{ text: root.textContent ?? '' }]
}

function getPPTFallbackHTMLParagraphAttributes(
  element: HTMLElement,
): Omit<PPTParagraph, 'runs'> {
  const style = element.getAttribute('style') ?? ''
  const align = parsePPTFallbackHTMLTextAlign(
    getPPTFallbackHTMLStyleValue(style, 'text-align'),
  )
  const lineHeight = parsePPTFallbackHTMLLineHeight(
    getPPTFallbackHTMLStyleValue(style, 'line-height'),
  )
  const spacingAfter = parsePPTFallbackHTMLPositivePixelStyle(
    style,
    'margin-bottom',
  )
  const spacingBefore = parsePPTFallbackHTMLPositivePixelStyle(
    style,
    'margin-top',
  )

  return {
    ...(align ? { align } : {}),
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
    ...(spacingBefore === undefined ? {} : { spacingBefore }),
  }
}

function collectPPTFallbackHTMLRuns(
  node: Node,
  inherited: Omit<PPTRun, 'text'>,
  runs: PPTRun[],
) {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? ''

    if (text) {
      runs.push({
        ...inherited,
        text,
      })
    }

    return
  }

  if (!(node instanceof HTMLElement)) {
    return
  }

  const tag = node.tagName.toLowerCase()
  const style = node.getAttribute('style') ?? ''
  const nextRunStyle = {
    ...inherited,
    ...(tag === 'strong' || tag === 'b'
      ? { bold: true }
      : {}),
    ...(tag === 'em' || tag === 'i'
      ? { italic: true }
      : {}),
    ...(tag === 'u'
      ? { underline: true }
      : {}),
    ...(getPPTFallbackHTMLStyleValue(style, 'font-weight') === '700' ||
      getPPTFallbackHTMLStyleValue(style, 'font-weight') === 'bold'
      ? { bold: true }
      : {}),
    ...(getPPTFallbackHTMLStyleValue(style, 'font-style') === 'italic'
      ? { italic: true }
      : {}),
    ...(getPPTFallbackHTMLStyleValue(style, 'text-decoration').includes('underline')
      ? { underline: true }
      : {}),
  }
  const color = parsePPTFallbackHTMLColor(
    getPPTFallbackHTMLStyleValue(style, 'color'),
    '',
  )
  const size = parsePPTFallbackHTMLPositivePixelStyle(style, 'font-size')
  const styledRun = {
    ...nextRunStyle,
    ...(color ? { color } : {}),
    ...(size === undefined ? {} : { size }),
  }

  for (const child of [...node.childNodes]) {
    collectPPTFallbackHTMLRuns(
      child,
      styledRun,
      runs,
    )
  }
}

function parsePPTFallbackHTMLTextAlign(
  value: string,
): PPTParagraph['align'] | undefined {
  return value === 'center' || value === 'right' || value === 'left'
    ? value
    : undefined
}

function parsePPTFallbackHTMLLineHeight(value: string) {
  const match = value.match(/-?\d+(?:\.\d+)?/)

  if (!match) {
    return undefined
  }

  const parsed = Number(match[0])

  if (!Number.isFinite(parsed) || parsed <= 0) {
    return undefined
  }

  return value.includes('%') ? parsed / 100 : parsed
}

function parsePPTFallbackHTMLPositivePixelStyle(
  style: string,
  property: string,
) {
  const value = parsePPTFallbackHTMLPixelStyle(style, property)

  return value !== undefined && value > 0 ? value : undefined
}
