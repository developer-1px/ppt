import type { CSSProperties } from 'react'
import {
  createSlideEditObjectAccessibilityDescriptor,
  createSlideEditObjectCornerRadiusDescriptor,
  createSlideEditObjectFillOpacityDescriptor,
  createSlideEditObjectHyperlinkDescriptor,
  createSlideEditObjectOpacityDescriptor,
  createSlideEditObjectShadowDescriptor,
  createSlideEditObjectStrokeLineStyleDescriptor,
  getSlideEditColorWithAlphaCSS,
  getSlideEditObjectStrokeLineStyleBorderStyle,
  getSlideEditObjectStrokeLineStyleDashArray,
  getSlideEditObjectShadowFilter,
  isSlideEditObjectStrokeLineStyleValue,
  normalizeSlideEditObjectAltTextStorageValue,
  normalizeSlideEditObjectCornerRadius,
  normalizeSlideEditObjectFillOpacity,
  normalizeSlideEditObjectHyperlinkStorageUrl,
  normalizeSlideEditObjectOpacity,
  normalizeSlideEditObjectStrokeLineStyle,
  SLIDE_EDIT_OBJECT_STROKE_LINE_STYLE_OPTIONS,
  toSlideEditObjectCornerRadiusAttributeValue,
  toSlideEditObjectFillOpacityAttributeValue,
  toSlideEditObjectOpacityAttributeValue,
  type SlideEditObjectAccessibilityDescriptor,
  type SlideEditObjectCornerRadiusDescriptor,
  type SlideEditObjectFillOpacityDescriptor,
  type SlideEditObjectHyperlinkDescriptor,
  type SlideEditObjectOpacityDescriptor,
  type SlideEditObjectShadowDescriptor,
  type SlideEditObjectStrokeLineStyleDescriptor,
} from './pptSlideEditAffordanceAdapter'
import { clampPPTCanvasValue } from './pptCanvasCoreAdapter'
import type {
  PPTElement,
  PPTElementAccessibility,
  PPTElementHyperlink,
  PPTElementShadow,
  PPTFill,
  PPTShape,
  PPTShapeKind,
  PPTStroke,
  PPTStrokeDash,
} from './pptModel'

export type PPTElementShadowUpdateField = keyof PPTElementShadow | 'enabled'

export const PPT_ELEMENT_OPACITY_MIN = 0
export const PPT_ELEMENT_OPACITY_MAX = 1
export const PPT_ELEMENT_OPACITY_STEP = 0.05
export const PPT_DEFAULT_ELEMENT_SHADOW = Object.freeze({
  angle: 45,
  blur: 14,
  color: '#000000',
  distance: 8,
  opacity: 0.22,
} as const satisfies PPTElementShadow)
export const PPT_ELEMENT_SHADOW_ANGLE_MIN = 0
export const PPT_ELEMENT_SHADOW_ANGLE_MAX = 359
export const PPT_ELEMENT_SHADOW_BLUR_MAX = 80
export const PPT_ELEMENT_SHADOW_DISTANCE_MAX = 120
export const PPT_ELEMENT_SHADOW_OPACITY_MIN = 0
export const PPT_ELEMENT_SHADOW_OPACITY_MAX = 1
export const PPT_ELEMENT_SHADOW_OPACITY_STEP = 0.05
export const PPT_ALT_TEXT_MAX_LENGTH = 1000
export const PPT_HYPERLINK_URL_MAX_LENGTH = 2048
export const PPT_FILL_OPACITY_MIN = 0
export const PPT_FILL_OPACITY_MAX = 1
export const PPT_FILL_OPACITY_STEP = 0.05
export const PPT_SHAPE_CORNER_RADIUS_DEFAULT = 24
export const PPT_SHAPE_CORNER_RADIUS_MIN = 0
export const PPT_SHAPE_CORNER_RADIUS_MAX = 120
export const PPT_SHAPE_CORNER_RADIUS_STEP = 1
export const PPT_STROKE_DASH_OPTIONS = Object.freeze(
  SLIDE_EDIT_OBJECT_STROKE_LINE_STYLE_OPTIONS.map((option) => ({
    id: option.id,
    label: option.label,
    value: option.id,
  })),
) as readonly {
  id: PPTStrokeDash
  label: string
  value: PPTStrokeDash
}[]

export function getPPTElementOpacity(element: PPTElement) {
  return normalizePPTElementOpacity(element.opacity ?? 1)
}

export function parsePPTElementOpacity(value: string) {
  return normalizePPTElementOpacity(Number(value))
}

export function normalizePPTElementOpacity(value: number) {
  return normalizeSlideEditObjectOpacity(value)
}

export function formatPPTElementOpacity(value: number) {
  return toSlideEditObjectOpacityAttributeValue(value)
}

export function getPPTObjectOpacityDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectOpacityDescriptor<string, string> {
  return createSlideEditObjectOpacityDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTElementOpacity(element),
  })
}

export function getPPTElementHyperlink(element: PPTElement) {
  return element.hyperlink ? normalizePPTElementHyperlink(element.hyperlink) : null
}

export function getPPTObjectHyperlinkDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectHyperlinkDescriptor<string, string> {
  const hyperlink = getPPTElementHyperlink(element)

  return createSlideEditObjectHyperlinkDescriptor({
    hyperlink: hyperlink
      ? {
          target: 'same-context',
          title: '',
          url: hyperlink.url,
        }
      : null,
    objectId: element.id,
    slideId,
  })
}

export function getPPTElementAltText(element: PPTElement) {
  return element.accessibility
    ? normalizePPTElementAccessibility(element.accessibility)?.altText
    : undefined
}

export function getPPTObjectAccessibilityDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectAccessibilityDescriptor<string, string> {
  return createSlideEditObjectAccessibilityDescriptor({
    objectId: element.id,
    slideId,
    value: {
      altText: getPPTElementAltText(element) ?? '',
      decorative: false,
    },
  })
}

export function normalizePPTElementAccessibility(
  accessibility: Partial<PPTElementAccessibility>,
): PPTElementAccessibility | null {
  const altText = normalizePPTAltText(accessibility.altText ?? '')

  return altText ? { altText } : null
}

export function normalizePPTAltText(value: string) {
  return normalizeSlideEditObjectAltTextStorageValue(value, {
    maxLength: PPT_ALT_TEXT_MAX_LENGTH,
  }) ?? ''
}

export function normalizePPTElementHyperlink(
  hyperlink: Partial<PPTElementHyperlink>,
): PPTElementHyperlink | null {
  const url = normalizePPTElementHyperlinkUrl(hyperlink.url ?? '')

  return url ? { url } : null
}

export function normalizePPTElementHyperlinkUrl(url: string) {
  return normalizeSlideEditObjectHyperlinkStorageUrl(url, {
    blockedSchemes: ['javascript', 'data', 'vbscript'],
    maxLength: PPT_HYPERLINK_URL_MAX_LENGTH,
  }) ?? ''
}

export function hasPPTElementShadow(element: PPTElement) {
  return element.shadow !== undefined
}

export function getPPTObjectShadowDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectShadowDescriptor<string, string> {
  return createSlideEditObjectShadowDescriptor({
    objectId: element.id,
    shadow: {
      ...getPPTElementShadow(element),
      enabled: hasPPTElementShadow(element),
    },
    slideId,
  })
}

export function getPPTObjectShadowField(
  descriptor: SlideEditObjectShadowDescriptor<string, string> | null,
  fieldId: PPTElementShadowUpdateField,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

export function getPPTElementShadow(element: PPTElement): PPTElementShadow {
  return normalizePPTElementShadow(element.shadow ?? PPT_DEFAULT_ELEMENT_SHADOW)
}

export function normalizePPTElementShadow(
  shadow: Partial<PPTElementShadow>,
): PPTElementShadow {
  return {
    angle: normalizePPTElementShadowAngle(
      shadow.angle ?? PPT_DEFAULT_ELEMENT_SHADOW.angle,
    ),
    blur: normalizePPTElementShadowBlur(
      shadow.blur ?? PPT_DEFAULT_ELEMENT_SHADOW.blur,
    ),
    color: normalizePPTElementShadowColor(
      shadow.color ?? PPT_DEFAULT_ELEMENT_SHADOW.color,
    ),
    distance: normalizePPTElementShadowDistance(
      shadow.distance ?? PPT_DEFAULT_ELEMENT_SHADOW.distance,
    ),
    opacity: normalizePPTElementShadowOpacity(
      shadow.opacity ?? PPT_DEFAULT_ELEMENT_SHADOW.opacity,
    ),
  }
}

export function parsePPTElementShadowAngle(value: string) {
  return normalizePPTElementShadowAngle(Number(value))
}

export function normalizePPTElementShadowAngle(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_DEFAULT_ELEMENT_SHADOW.angle

  return clampPPTCanvasValue(
    Math.round(finiteValue),
    PPT_ELEMENT_SHADOW_ANGLE_MIN,
    PPT_ELEMENT_SHADOW_ANGLE_MAX,
  )
}

export function parsePPTElementShadowBlur(value: string) {
  return normalizePPTElementShadowBlur(Number(value))
}

export function normalizePPTElementShadowBlur(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_DEFAULT_ELEMENT_SHADOW.blur

  return clampPPTCanvasValue(
    Math.round(finiteValue),
    0,
    PPT_ELEMENT_SHADOW_BLUR_MAX,
  )
}

export function parsePPTElementShadowDistance(value: string) {
  return normalizePPTElementShadowDistance(Number(value))
}

export function normalizePPTElementShadowDistance(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_DEFAULT_ELEMENT_SHADOW.distance

  return clampPPTCanvasValue(
    Math.round(finiteValue),
    0,
    PPT_ELEMENT_SHADOW_DISTANCE_MAX,
  )
}

export function parsePPTElementShadowOpacity(value: string) {
  return normalizePPTElementShadowOpacity(Number(value))
}

export function normalizePPTElementShadowOpacity(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_DEFAULT_ELEMENT_SHADOW.opacity
  const clamped = clampPPTCanvasValue(
    finiteValue,
    PPT_ELEMENT_SHADOW_OPACITY_MIN,
    PPT_ELEMENT_SHADOW_OPACITY_MAX,
  )

  return Math.round(clamped * 100) / 100
}

export function formatPPTElementShadowOpacity(value: number) {
  return toSlideEditObjectOpacityAttributeValue(value)
}

export function normalizePPTElementShadowColor(color: string) {
  return /^#[\da-f]{6}$/i.test(color)
    ? color
    : PPT_DEFAULT_ELEMENT_SHADOW.color
}

export function getPPTElementShadowFilter(element: PPTElement) {
  return getSlideEditObjectShadowFilter({
    ...getPPTElementShadow(element),
    enabled: hasPPTElementShadow(element),
  })
}

export function getPPTElementBorderStyle(
  element: PPTElement,
): CSSProperties {
  const stroke = getPPTElementStroke(element)

  if (!stroke || (element.kind !== 'shape' && element.kind !== 'image')) {
    return {}
  }

  return {
    border: getPPTStrokeBorderCSS(stroke),
    borderStyle: getPPTStrokeDashBorderStyle(stroke),
  }
}

function getPPTStrokeBorderCSS(stroke: PPTStroke) {
  return `${stroke.width}px solid ${stroke.color}`
}

export function getPPTElementStrokeDash(element: PPTElement) {
  const stroke = getPPTElementStroke(element)

  return stroke ? getPPTStrokeDash(stroke) : undefined
}

export function getPPTShapeCornerRadius(element: PPTShape) {
  return element.shape === 'rect'
    ? normalizePPTShapeCornerRadius(
        element.cornerRadius ?? PPT_SHAPE_CORNER_RADIUS_DEFAULT,
      )
    : 0
}

export function getPPTShapeCornerRadiusModelValue(value: number) {
  const normalized = normalizePPTShapeCornerRadius(value)

  return normalized === PPT_SHAPE_CORNER_RADIUS_DEFAULT
    ? undefined
    : normalized
}

export function parsePPTShapeCornerRadius(value: string) {
  return normalizePPTShapeCornerRadius(Number(value))
}

export function normalizePPTShapeCornerRadius(value: number) {
  return normalizeSlideEditObjectCornerRadius(value)
}

export function formatPPTShapeCornerRadius(value: number) {
  return toSlideEditObjectCornerRadiusAttributeValue(value)
}

export function getPPTCornerRadiusDescriptor(
  slideId: string,
  element: PPTShape,
): SlideEditObjectCornerRadiusDescriptor<string, string> {
  const isSupported = element.shape === 'rect'

  return createSlideEditObjectCornerRadiusDescriptor({
    isSupported,
    objectId: element.id,
    slideId,
    unsupportedReason: isSupported ? undefined : 'unsupported-shape',
    value: getPPTShapeCornerRadius(element),
  })
}

export function getPPTFillOpacityDescriptor(
  slideId: string,
  element: PPTShape,
): SlideEditObjectFillOpacityDescriptor<string, string> {
  return createSlideEditObjectFillOpacityDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTFillOpacity(element.fill),
  })
}

export function normalizePPTFill(fill: Partial<PPTFill>): PPTFill {
  const opacity = normalizePPTFillOpacity(fill.opacity ?? 1)
  const normalized = {
    color: typeof fill.color === 'string' && fill.color
      ? fill.color
      : '#ffffff',
  }

  return opacity === 1
    ? normalized
    : { ...normalized, opacity }
}

export function getPPTFillOpacity(fill: PPTFill) {
  return normalizePPTFillOpacity(fill.opacity ?? 1)
}

export function parsePPTFillOpacity(value: string) {
  return normalizePPTFillOpacity(Number(value))
}

export function normalizePPTFillOpacity(value: number) {
  return normalizeSlideEditObjectFillOpacity(value)
}

export function formatPPTFillOpacity(value: number) {
  return toSlideEditObjectFillOpacityAttributeValue(value)
}

export function getPPTFillColorCSS(fill: PPTFill) {
  const opacity = getPPTFillOpacity(fill)

  if (opacity === 1) {
    return fill.color
  }

  return getSlideEditColorWithAlphaCSS({
    color: fill.color,
    opacity,
  })
}

export function getPPTElementStroke(element: PPTElement): PPTStroke | null {
  if (element.kind === 'shape' || element.kind === 'image') {
    return element.stroke ? normalizePPTStroke(element.stroke) : null
  }

  if (element.kind === 'line' || element.kind === 'freeform') {
    return normalizePPTStroke(element.stroke)
  }

  return null
}

export function getPPTStrokeLineStyleDescriptor(
  slideId: string,
  element: PPTElement,
): SlideEditObjectStrokeLineStyleDescriptor<string, string> {
  const stroke = getPPTElementStroke(element)

  return createSlideEditObjectStrokeLineStyleDescriptor({
    isSupported: Boolean(stroke),
    objectId: element.id,
    slideId,
    unsupportedReason: stroke ? undefined : 'no-stroke',
    value: getPPTStrokeDash(stroke ?? undefined),
  })
}

export function normalizePPTStroke(stroke: Partial<PPTStroke>): PPTStroke {
  const dash = normalizePPTStrokeDash(stroke.dash)
  const normalized = {
    color: typeof stroke.color === 'string' && stroke.color
      ? stroke.color
      : '#111827',
    width: normalizePPTStrokeWidth(stroke.width ?? 2),
  }

  return dash === 'solid'
    ? normalized
    : { ...normalized, dash }
}

export function normalizePPTStrokeWidth(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 2

  return Math.max(0, Math.min(40, finiteValue))
}

export function getPPTStrokeDash(
  stroke: PPTStroke | undefined,
): PPTStrokeDash {
  return normalizePPTStrokeDash(stroke?.dash)
}

export function normalizePPTStrokeDash(value: unknown): PPTStrokeDash {
  return normalizeSlideEditObjectStrokeLineStyle(
    typeof value === 'string' ? value : null,
  ) as PPTStrokeDash
}

export function isPPTStrokeDash(value: string): value is PPTStrokeDash {
  return isSlideEditObjectStrokeLineStyleValue(value)
}

export function getPPTStrokeDashBorderStyle(
  stroke: PPTStroke | undefined,
) {
  return getSlideEditObjectStrokeLineStyleBorderStyle(getPPTStrokeDash(stroke))
}

export function getPPTStrokeDashArray(stroke: PPTStroke | undefined) {
  return getSlideEditObjectStrokeLineStyleDashArray({
    strokeWidth: normalizePPTStrokeWidth(stroke?.width ?? 2),
    value: getPPTStrokeDash(stroke),
  })
}

export function getPPTThumbLineDashStyle(
  element: PPTElement,
): CSSProperties {
  if (element.kind !== 'line') {
    return {}
  }

  return {
    '--ppt-thumb-line-dash': getPPTStrokeDashBorderStyle(element.stroke),
  } as CSSProperties
}

export function getPPTShapeLabel(shape: PPTShapeKind) {
  if (shape === 'ellipse') {
    return 'Oval'
  }

  if (shape === 'diamond') {
    return 'Diamond'
  }

  return 'Rectangle'
}

export function isPPTShapeKind(value: string): value is PPTShapeKind {
  return value === 'rect' || value === 'ellipse' || value === 'diamond'
}
