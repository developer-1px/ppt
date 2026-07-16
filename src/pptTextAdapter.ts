import type { CSSProperties } from 'react'
import {
  createSlideEditTextFontFamilyDescriptor,
  createSlideEditTextFrameInsetDescriptor,
  createSlideEditTextParagraphListLevelDescriptor,
  createSlideEditTextParagraphSpacingDescriptor,
  createSlideEditTextVerticalAlignmentDescriptor,
  getSlideEditTextFontFamilyCSS,
  getSlideEditTextParagraphListLevelIndentCSSValue,
  getSlideEditTextParagraphListLevelModelValue,
  getSlideEditTextParagraphSpacingCSSStyle,
  getSlideEditTextVerticalAlignmentFlexAlignItems,
  normalizeSlideEditTextFontFamily,
  normalizeSlideEditTextFrameInsetValue,
  normalizeSlideEditTextLineHeightRatio,
  normalizeSlideEditTextParagraphListLevel,
  normalizeSlideEditTextParagraphSpacingAmount,
  normalizeSlideEditTextVerticalAlignment,
  SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_LIMITS,
  SLIDE_EDIT_TEXT_VERTICAL_ALIGNMENT_OPTIONS,
  type SlideEditTextBoxSizeMode,
  type SlideEditTextFontFamilyDescriptor,
  type SlideEditTextFrameInsetDescriptor,
  type SlideEditTextParagraphListLevelDescriptor,
  type SlideEditTextParagraphSpacingDescriptor,
  type SlideEditTextParagraphSpacingFieldId,
  type SlideEditTextParagraphSpacingUpdateCommand,
  type SlideEditTextVerticalAlignmentDescriptor,
} from './pptSlideEditAffordanceAdapter'
import type {
  PPTElement,
  PPTParagraph,
  PPTTextAutoFit,
  PPTTextBody,
  PPTTextElement,
  PPTTextStyle,
} from './pptModel'

export type PPTParagraphSpacingField =
  | 'lineHeight'
  | 'spacingAfter'
  | 'spacingBefore'

export type PPTTextInset = NonNullable<PPTTextStyle['textInset']>
export type PPTTextInsetField = keyof PPTTextInset
export type PPTTextVerticalAlign = NonNullable<PPTTextStyle['verticalAlign']>

type PPTParagraphTextAlign =
  | 'center'
  | 'justify'
  | 'left'
  | 'right'
  | 'start'
  | 'end'
  | 'match-parent'

type PPTParagraphCSSStyle = CSSProperties & ReturnType<
  typeof getSlideEditTextParagraphSpacingCSSStyle
> & {
  '--ppt-paragraph-list-level-indent'?: string
  textAlign?: PPTParagraphTextAlign
}

export const PPT_DEFAULT_TEXT_FONT_FAMILY = 'Inter'
export const PPT_TEXT_FONT_FAMILY_OPTIONS = Object.freeze([
  {
    css: 'Inter, ui-sans-serif, system-ui, sans-serif',
    label: 'Inter',
    value: 'Inter',
  },
  { css: 'Arial, Helvetica, sans-serif', label: 'Arial', value: 'Arial' },
  { css: 'Georgia, serif', label: 'Georgia', value: 'Georgia' },
  {
    css: '"Courier New", monospace',
    label: 'Courier New',
    value: 'Courier New',
  },
] as const)
export const PPT_DEFAULT_TEXT_VERTICAL_ALIGN: PPTTextVerticalAlign = 'top'
export const PPT_TEXT_VERTICAL_ALIGNMENT_OPTIONS =
  SLIDE_EDIT_TEXT_VERTICAL_ALIGNMENT_OPTIONS
export const PPT_TEXT_INSET_MIN = 0
export const PPT_TEXT_INSET_MAX = 120
export const PPT_TEXT_INSET_STEP = 2
export const PPT_DEFAULT_TEXT_BOX_INSET = Object.freeze({
  bottom: 0,
  left: 0,
  right: 0,
  top: 0,
} as const satisfies PPTTextInset)
export const PPT_DEFAULT_SHAPE_TEXT_INSET = Object.freeze({
  bottom: 18,
  left: 18,
  right: 18,
  top: 18,
} as const satisfies PPTTextInset)
export const PPT_TEXT_AUTOFIT: PPTTextAutoFit = 'resizeShapeToFitText'
export const PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT = 1.14
export const PPT_PARAGRAPH_LINE_HEIGHT_MIN = 0.8
export const PPT_PARAGRAPH_LINE_HEIGHT_MAX = 3
export const PPT_PARAGRAPH_SPACING_MAX = 240
export const PPT_PARAGRAPH_LIST_LEVEL_MIN =
  SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_LIMITS.min
export const PPT_PARAGRAPH_LIST_LEVEL_MAX =
  SLIDE_EDIT_TEXT_PARAGRAPH_LIST_LEVEL_LIMITS.max

function getDefaultPPTParagraphSpacing() {
  return {
    lineHeight: PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
    spacingAfter: 0,
    spacingBefore: 0,
  }
}

export function getPPTTextElementParagraphSpacing(element: PPTTextElement) {
  const paragraph = element.textBody.paragraphs[0]

  if (!paragraph) {
    return getDefaultPPTParagraphSpacing()
  }

  return {
    lineHeight: getPPTParagraphLineHeight(paragraph),
    spacingAfter: getPPTParagraphSpacingAfter(paragraph),
    spacingBefore: getPPTParagraphSpacingBefore(paragraph),
  }
}

export function getPPTTextParagraphSpacingDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextParagraphSpacingDescriptor<string, string> {
  const spacing = getPPTTextElementParagraphSpacing(element)

  return createSlideEditTextParagraphSpacingDescriptor({
    lineHeightRatio: spacing.lineHeight,
    objectId: element.id,
    paragraphAfter: {
      unit: 'px',
      value: spacing.spacingAfter,
    },
    paragraphBefore: {
      unit: 'px',
      value: spacing.spacingBefore,
    },
    slideId,
  })
}

export function getPPTTextParagraphListLevelDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextParagraphListLevelDescriptor<string, string> {
  return createSlideEditTextParagraphListLevelDescriptor({
    level: getPPTParagraphListLevel(
      element.textBody.paragraphs[0] ?? { runs: [] },
    ),
    objectId: element.id,
    slideId,
  })
}

export function getPPTTextParagraphSpacingField(
  descriptor: SlideEditTextParagraphSpacingDescriptor<string, string> | null,
  fieldId: SlideEditTextParagraphSpacingFieldId,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

export function toSlideEditParagraphSpacingCommand({
  elementId,
  field,
  slideId,
  value,
}: {
  elementId: string
  field: PPTParagraphSpacingField
  slideId: string
  value: number
}): SlideEditTextParagraphSpacingUpdateCommand<string, string> {
  if (field === 'lineHeight') {
    return {
      fieldId: 'lineHeightRatio',
      id: 'update-text-paragraph-spacing',
      objectId: elementId,
      slideId,
      value,
    }
  }

  return {
    fieldId: field === 'spacingBefore' ? 'paragraphBefore' : 'paragraphAfter',
    id: 'update-text-paragraph-spacing',
    objectId: elementId,
    slideId,
    value: {
      unit: 'px',
      value,
    },
  }
}

export function toPPTParagraphSpacingUpdate(
  command: SlideEditTextParagraphSpacingUpdateCommand<string, string>,
): {
  field: PPTParagraphSpacingField
  value: number
} {
  if (command.fieldId === 'lineHeightRatio') {
    return {
      field: 'lineHeight',
      value: normalizePPTParagraphLineHeight(command.value),
    }
  }

  return {
    field: command.fieldId === 'paragraphBefore'
      ? 'spacingBefore'
      : 'spacingAfter',
    value: normalizePPTParagraphSpacing(command.value.value),
  }
}

export function normalizePPTParagraphLineHeight(value: number) {
  const finiteValue = Number.isFinite(value)
    ? value
    : PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT

  return normalizeSlideEditTextLineHeightRatio(finiteValue)
}

export function normalizePPTParagraphSpacing(value: number) {
  return normalizeSlideEditTextParagraphSpacingAmount({
    unit: 'px',
    value: Number.isFinite(value) ? value : 0,
  }).value
}

export function parsePPTParagraphLineHeight(value: string) {
  return normalizePPTParagraphLineHeight(Number(value))
}

export function parsePPTParagraphSpacing(value: string) {
  return normalizePPTParagraphSpacing(Number(value))
}

export function normalizePPTParagraphListLevel(
  value: number | null | undefined,
) {
  return normalizeSlideEditTextParagraphListLevel(value)
}

export function getPPTParagraphListLevel(paragraph: PPTParagraph) {
  return normalizePPTParagraphListLevel(paragraph.level)
}

export function getPPTParagraphListLevelModelValue(value: number) {
  return getSlideEditTextParagraphListLevelModelValue(value)
}

export function getPPTParagraphLineHeight(paragraph: PPTParagraph) {
  return normalizePPTParagraphLineHeight(
    paragraph.lineHeight ?? PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
  )
}

export function getPPTParagraphSpacingAfter(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingAfter ?? 0)
}

export function getPPTParagraphSpacingBefore(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingBefore ?? 0)
}

export function getPPTParagraphStyle(
  paragraph: PPTParagraph,
): PPTParagraphCSSStyle {
  const listLevel = getPPTParagraphListLevel(paragraph)

  return {
    ...getSlideEditTextParagraphSpacingCSSStyle({
      lineHeightRatio: getPPTParagraphLineHeight(paragraph),
      paragraphAfter: {
        unit: 'px',
        value: getPPTParagraphSpacingAfter(paragraph),
      },
      paragraphBefore: {
        unit: 'px',
        value: getPPTParagraphSpacingBefore(paragraph),
      },
    }),
    '--ppt-paragraph-list-level-indent':
      getSlideEditTextParagraphListLevelIndentCSSValue(listLevel),
    textAlign: paragraph.align,
  }
}

export function getDefaultPPTTextStyle(): PPTTextStyle {
  return {
    color: '#111827',
    fontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontSize: 24,
    fontWeight: 'regular',
    verticalAlign: PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
  }
}

export function getPPTTextElementStyle(
  element: PPTTextElement,
): PPTTextStyle {
  return {
    ...getDefaultPPTTextStyle(),
    ...element.style,
  }
}

export function hasPPTTextBodyBullet(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'bullet')
}

export function hasPPTTextBodyNumbered(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'numbered')
}

export function getPPTTextFontFamilyDescriptorOptions() {
  return PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => ({
    cssFontFamily: option.css,
    family: option.value,
    isDefault: option.value === PPT_DEFAULT_TEXT_FONT_FAMILY,
    label: option.label,
    source: 'host' as const,
  }))
}

export function getPPTTextFontFamilyDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextFontFamilyDescriptor<string, string> {
  return createSlideEditTextFontFamilyDescriptor({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily: getPPTTextElementStyle(element).fontFamily,
    objectId: element.id,
    options: getPPTTextFontFamilyDescriptorOptions(),
    slideId,
  })
}

export function normalizePPTTextFontFamily(fontFamily: string | undefined) {
  return normalizeSlideEditTextFontFamily({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily,
    options: getPPTTextFontFamilyDescriptorOptions(),
  })
}

export function getPPTTextFontFamilyCSS(fontFamily: string | undefined) {
  return getSlideEditTextFontFamilyCSS({
    fallbackFontFamily: PPT_DEFAULT_TEXT_FONT_FAMILY,
    fontFamily,
    options: getPPTTextFontFamilyDescriptorOptions(),
  })
}

export function getPPTTextElementVerticalAlign(element: PPTElement) {
  const verticalAlign =
    element.kind === 'freeform' ||
    element.kind === 'shape' ||
    element.kind === 'textBox'
      ? element.style?.verticalAlign
      : undefined

  return normalizePPTTextVerticalAlign(
    verticalAlign,
    element.kind === 'freeform' || element.kind === 'shape'
      ? 'middle'
      : PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
  )
}

export function getPPTTextVerticalAlignmentDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextVerticalAlignmentDescriptor<string, string> {
  return createSlideEditTextVerticalAlignmentDescriptor({
    objectId: element.id,
    slideId,
    value: getPPTTextElementVerticalAlign(element),
  })
}

export function normalizePPTTextVerticalAlign(
  verticalAlign: string | undefined,
  fallback: PPTTextVerticalAlign = PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
) {
  const normalizedFallback = normalizeSlideEditTextVerticalAlignment(fallback)
  const normalizedValue = normalizeSlideEditTextVerticalAlignment(verticalAlign)

  return verticalAlign === undefined ? normalizedFallback : normalizedValue
}

export function getPPTTextVerticalAlignCSS(
  verticalAlign: string | undefined,
) {
  return getSlideEditTextVerticalAlignmentFlexAlignItems(
    normalizePPTTextVerticalAlign(verticalAlign),
  )
}

export function getPPTTextFrameInsetDescriptor(
  slideId: string,
  element: PPTTextElement,
): SlideEditTextFrameInsetDescriptor<string, string> {
  return createSlideEditTextFrameInsetDescriptor({
    inset: getPPTTextElementInset(element),
    objectId: element.id,
    slideId,
  })
}

export function getPPTTextFrameInsetField(
  descriptor: SlideEditTextFrameInsetDescriptor<string, string> | null,
  fieldId: PPTTextInsetField,
) {
  return descriptor?.fields.find((field) => field.id === fieldId)
}

export function getPPTTextElementInset(element: PPTElement): PPTTextInset {
  const fallback = element.kind === 'freeform' || element.kind === 'shape'
    ? PPT_DEFAULT_SHAPE_TEXT_INSET
    : PPT_DEFAULT_TEXT_BOX_INSET
  const inset =
    element.kind === 'freeform' ||
    element.kind === 'shape' ||
    element.kind === 'textBox'
      ? element.style?.textInset
      : undefined

  return {
    bottom: normalizePPTTextInset(inset?.bottom ?? fallback.bottom),
    left: normalizePPTTextInset(inset?.left ?? fallback.left),
    right: normalizePPTTextInset(inset?.right ?? fallback.right),
    top: normalizePPTTextInset(inset?.top ?? fallback.top),
  }
}

export function parsePPTTextInset(value: string) {
  return normalizePPTTextInset(Number(value))
}

export function normalizePPTTextInset(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 0

  return normalizeSlideEditTextFrameInsetValue(finiteValue)
}

export function formatPPTTextInsetData(inset: PPTTextInset) {
  return `${inset.top},${inset.right},${inset.bottom},${inset.left}`
}

export function getPPTTextAutoFitSizeMode(
  element: PPTTextElement,
): SlideEditTextBoxSizeMode {
  return element.textAutoFit === PPT_TEXT_AUTOFIT ? 'resize-to-fit' : 'fixed'
}
