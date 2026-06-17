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
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTShapeKind,
  type PPTStroke,
  type PPTTextBody,
  type PPTTextStyle,
} from '../pptModel'

export const PPT_FALLBACK_HTML_IMPORT_MODEL = 'ppt-fallback-html-import'

export type PPTFallbackHTMLShapeSource = {
  cornerRadius?: number
  fill: PPTFill
  geometry: {
    h: number
    w: number
  }
  name: string
  shape: PPTShapeKind
  sourceObjectId?: string
  stroke?: PPTStroke
  style: PPTTextStyle
  textBody?: PPTTextBody
}

export type PPTFallbackHTMLImportEffect = {
  format: 'text-html-ppt-fallback'
  model: typeof PPT_FALLBACK_HTML_IMPORT_MODEL
  name: string
  shape: PPTShapeKind
  sourceObjectId?: string
}

const PPT_FALLBACK_SHAPE_MIN_SIZE = 48
const PPT_FALLBACK_SHAPE_MAX_WIDTH = 980
const PPT_FALLBACK_SHAPE_MAX_HEIGHT = 560

export function getPPTFallbackHTMLShapeSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTFallbackHTMLShapeSourceFromHTML(
    dataTransfer?.getData('text/html') ?? '',
  )
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
      textInset: parsePPTFallbackHTMLPadding(rawStyle),
      verticalAlign: parsePPTFallbackHTMLVerticalAlign(
        getPPTFallbackHTMLStyleValue(rawStyle, 'align-items'),
      ),
    },
    ...(textBody ? { textBody } : {}),
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

export function createPPTFallbackHTMLImportEffect({
  element,
  source,
}: {
  element: PPTShape
  source: PPTFallbackHTMLShapeSource
}): PPTFallbackHTMLImportEffect {
  return {
    format: 'text-html-ppt-fallback',
    model: PPT_FALLBACK_HTML_IMPORT_MODEL,
    name: element.name,
    shape: source.shape,
    ...(source.sourceObjectId ? { sourceObjectId: source.sourceObjectId } : {}),
  }
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

function parsePPTFallbackHTMLPadding(style: string) {
  const value = getPPTFallbackHTMLStyleValue(style, 'padding')
  const parts = value
    .split(/\s+/)
    .map((part) => Number(part.match(/-?\d+(?:\.\d+)?/)?.[0] ?? NaN))
    .filter((part) => Number.isFinite(part))

  if (parts.length === 0) {
    return {
      bottom: 18,
      left: 18,
      right: 18,
      top: 18,
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
        ...(paragraphElement.tagName.toLowerCase() === 'li'
          ? { bullet: 'bullet' as const }
          : {}),
        runs,
      }
    })
    .filter((paragraph): paragraph is PPTParagraph => paragraph !== null)

  return paragraphs.length > 0 ? { paragraphs } : null
}

function getPPTFallbackHTMLRuns(root: HTMLElement): PPTRun[] {
  const runs: PPTRun[] = []

  collectPPTFallbackHTMLRuns(root, {}, runs)

  return runs.length > 0 ? runs : [{ text: root.textContent ?? '' }]
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

  for (const child of [...node.childNodes]) {
    collectPPTFallbackHTMLRuns(
      child,
      color ? { ...nextRunStyle, color } : nextRunStyle,
      runs,
    )
  }
}
