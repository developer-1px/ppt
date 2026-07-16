import {
  createSlideEditObjectAccessibilityDescriptor,
  createSlideEditObjectHyperlinkDescriptor,
  createSlideEditObjectOpacityDescriptor,
  createSlideEditObjectShadowDescriptor,
  getSlideEditObjectShadowFilter,
  normalizeSlideEditObjectAltTextStorageValue,
  normalizeSlideEditObjectHyperlinkStorageUrl,
  normalizeSlideEditObjectOpacity,
  toSlideEditObjectOpacityAttributeValue,
  type SlideEditObjectAccessibilityDescriptor,
  type SlideEditObjectHyperlinkDescriptor,
  type SlideEditObjectOpacityDescriptor,
  type SlideEditObjectShadowDescriptor,
} from './pptSlideEditAffordanceAdapter'
import { clampPPTCanvasValue } from './pptCanvasCoreAdapter'
import type {
  PPTElement,
  PPTElementAccessibility,
  PPTElementHyperlink,
  PPTElementShadow,
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
