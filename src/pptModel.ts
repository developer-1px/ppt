import { z } from 'zod'
import { createPPTCanvasSequentialIdFactory } from './pptCanvasCoreAdapter'

export const PPT_SLIDE_WIDTH = 1280
export const PPT_SLIDE_HEIGHT = 720
export const PPT_DEFAULT_THEME_ID = 'ppt-theme-default'
export const PPT_TITLE_BODY_LAYOUT_ID = 'ppt-layout-title-body'
export const PPT_SPLIT_LAYOUT_ID = 'ppt-layout-split'

const PPTFillSchema = z.object({
  color: z.string(),
  opacity: z.number().optional(),
})
const PPTStrokeSchema = z.object({
  color: z.string(),
  dash: z.enum(['solid', 'dash', 'dot']).optional(),
  width: z.number(),
})
const PPTTextInsetSchema = z.object({
  bottom: z.number(),
  left: z.number(),
  right: z.number(),
  top: z.number(),
})
const PPTTableCellBordersSchema = z.object({
  bottom: PPTStrokeSchema.optional(),
  left: PPTStrokeSchema.optional(),
  right: PPTStrokeSchema.optional(),
  top: PPTStrokeSchema.optional(),
})
const PPTTableCellTextStyleSchema = z.object({
  align: z.enum(['left', 'center', 'right', 'justify']).optional(),
  color: z.string().optional(),
  fontSize: z.number().optional(),
  fontWeight: z.enum(['regular', 'semibold', 'bold']).optional(),
  textInset: PPTTextInsetSchema.optional(),
  verticalAlign: z.enum(['top', 'middle', 'bottom']).optional(),
})
const PPTTableCellStyleSchema = z.object({
  borders: PPTTableCellBordersSchema.optional(),
  colSpan: z.number().int().positive().optional(),
  fill: PPTFillSchema.optional(),
  hidden: z.boolean().optional(),
  rowSpan: z.number().int().positive().optional(),
  textStyle: PPTTableCellTextStyleSchema.optional(),
})

const PPTGeometrySchema = z.object({
  h: z.number(),
  rotation: z.number().optional(),
  w: z.number(),
  x: z.number(),
  y: z.number(),
})

const PPTRunSchema = z.object({
  bold: z.boolean().optional(),
  color: z.string().optional(),
  highlight: z.string().optional(),
  italic: z.boolean().optional(),
  size: z.number().optional(),
  strikethrough: z.boolean().optional(),
  text: z.string(),
  underline: z.boolean().optional(),
})

const PPTParagraphSchema = z.object({
  align: z.enum(['left', 'center', 'right', 'justify']).optional(),
  bullet: z.enum(['bullet', 'numbered']).optional(),
  level: z.number().int().nonnegative().optional(),
  lineHeight: z.number().optional(),
  runs: z.array(PPTRunSchema),
  spacingAfter: z.number().optional(),
  spacingBefore: z.number().optional(),
})

export const PPTTextBodySchema = z.object({
  paragraphs: z.array(PPTParagraphSchema),
})

const PPTTextStyleSchema = z.object({
  color: z.string(),
  fontFamily: z.string().optional(),
  fontSize: z.number(),
  fontWeight: z.enum(['regular', 'semibold', 'bold']).optional(),
  textInset: PPTTextInsetSchema.optional(),
  verticalAlign: z.enum(['top', 'middle', 'bottom']).optional(),
})

const PPTTextAutoFitSchema = z.enum(['resizeShapeToFitText'])

const PPTSlideTransitionSchema = z.object({
  advanceAfterMs: z.number().nullable().optional(),
  advanceOnClick: z.boolean().optional(),
  durationMs: z.number(),
  type: z.enum(['none', 'fade', 'push']),
})

const PPTElementAnimationSchema = z.object({
  delayMs: z.number(),
  durationMs: z.number(),
  order: z.number(),
  trigger: z.enum(['onClick', 'withPrevious']),
  type: z.enum(['none', 'fadeIn', 'flyIn']),
})

const PPTElementShadowSchema = z.object({
  angle: z.number(),
  blur: z.number(),
  color: z.string(),
  distance: z.number(),
  opacity: z.number(),
})

const PPTElementHyperlinkSchema = z.object({
  url: z.string(),
})

const PPTElementAccessibilitySchema = z.object({
  altText: z.string(),
})

const PPTImageAdjustmentsSchema = z.object({
  brightness: z.number().optional(),
  contrast: z.number().optional(),
  grayscale: z.boolean().optional(),
})
const PPTImageClipShapeSchema = z.enum(['ellipse', 'diamond'])

const PPTElementBaseSchema = z.object({
  accessibility: PPTElementAccessibilitySchema.optional(),
  animation: PPTElementAnimationSchema.optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  geometry: PPTGeometrySchema,
  groupId: z.string().optional(),
  hyperlink: PPTElementHyperlinkSchema.optional(),
  id: z.string(),
  locked: z.boolean().optional(),
  name: z.string(),
  opacity: z.number().optional(),
  shadow: PPTElementShadowSchema.optional(),
  visible: z.boolean().optional(),
})

const PPTTextBoxSchema = PPTElementBaseSchema.extend({
  kind: z.literal('textBox'),
  style: PPTTextStyleSchema,
  textAutoFit: PPTTextAutoFitSchema.optional(),
  textBody: PPTTextBodySchema,
})

const PPTShapeSchema = PPTElementBaseSchema.extend({
  cornerRadius: z.number().optional(),
  fill: PPTFillSchema,
  kind: z.literal('shape'),
  shape: z.enum(['rect', 'ellipse', 'diamond']),
  stroke: PPTStrokeSchema.optional(),
  style: PPTTextStyleSchema.optional(),
  textAutoFit: PPTTextAutoFitSchema.optional(),
  textBody: PPTTextBodySchema.optional(),
})

const PPTImageSchema = PPTElementBaseSchema.extend({
  adjustments: PPTImageAdjustmentsSchema.optional(),
  alt: z.string(),
  clipShape: PPTImageClipShapeSchema.optional(),
  crop: z.object({
    bottom: z.number().optional(),
    left: z.number().optional(),
    right: z.number().optional(),
    top: z.number().optional(),
    x: z.number(),
    y: z.number(),
  }).optional(),
  fit: z.enum(['cover', 'contain']).optional(),
  kind: z.literal('image'),
  src: z.string(),
})

const PPTLineMarkerSchema = z.enum(['none', 'arrow'])
const PPTLineRouteSchema = z.enum(['straight', 'elbow'])

const PPTLinePointSchema = z.object({
  x: z.number(),
  y: z.number(),
})
const PPTFreeformPathSegmentSchema = z.discriminatedUnion('type', [
  z.object({
    point: PPTLinePointSchema,
    type: z.literal('move'),
  }),
  z.object({
    point: PPTLinePointSchema,
    type: z.literal('line'),
  }),
  z.object({
    control1: PPTLinePointSchema,
    control2: PPTLinePointSchema,
    point: PPTLinePointSchema,
    type: z.literal('cubic'),
  }),
])

const PPTLineConnectionSchema = z.object({
  anchor: z.enum(['bottom', 'center', 'left', 'right', 'top']),
  elementId: z.string(),
})

const PPTLineSchema = PPTElementBaseSchema.extend({
  end: PPTLinePointSchema,
  endConnection: PPTLineConnectionSchema.optional(),
  endMarker: PPTLineMarkerSchema.optional(),
  kind: z.literal('line'),
  route: PPTLineRouteSchema.optional(),
  routeBend: z.number().optional(),
  start: PPTLinePointSchema,
  startConnection: PPTLineConnectionSchema.optional(),
  startMarker: PPTLineMarkerSchema.optional(),
  stroke: PPTStrokeSchema,
})

const PPTFreeformSchema = PPTElementBaseSchema.extend({
  fill: PPTFillSchema.optional(),
  kind: z.literal('freeform'),
  pointMode: z.enum(['freehand', 'polyline']).optional(),
  points: z.array(PPTLinePointSchema).min(1),
  segments: z.array(PPTFreeformPathSegmentSchema).optional(),
  style: PPTTextStyleSchema.optional(),
  stroke: PPTStrokeSchema,
  textAutoFit: PPTTextAutoFitSchema.optional(),
  textBody: PPTTextBodySchema.optional(),
})

const PPTTableSchema = PPTElementBaseSchema.extend({
  cellStyles: z.array(z.array(PPTTableCellStyleSchema)).optional(),
  columnWidths: z.array(z.number()).optional(),
  kind: z.literal('table'),
  rowHeights: z.array(z.number()).optional(),
  rows: z.array(z.array(z.string())),
})

const PPTCommentThreadMessageSchema = z.object({
  authorName: z.string(),
  body: z.string(),
  createdAt: z.string(),
  id: z.string(),
})

const PPTCommentSchema = PPTElementBaseSchema.extend({
  authorName: z.string().optional(),
  body: z.string(),
  createdAt: z.string().optional(),
  kind: z.literal('comment'),
  resolved: z.boolean().optional(),
  thread: z.array(PPTCommentThreadMessageSchema).optional(),
})

export const PPTElementSchema = z.discriminatedUnion('kind', [
  PPTTextBoxSchema,
  PPTShapeSchema,
  PPTImageSchema,
  PPTLineSchema,
  PPTFreeformSchema,
  PPTTableSchema,
  PPTCommentSchema,
])

export const PPTSlideSchema = z.object({
  background: PPTFillSchema.optional(),
  elements: z.array(PPTElementSchema),
  hidden: z.boolean().optional(),
  hiddenPlaceholderIds: z.array(z.string()).optional(),
  id: z.string(),
  layoutId: z.string().optional(),
  name: z.string(),
  notes: z.string().optional(),
  sectionName: z.string().optional(),
  themeId: z.string().optional(),
  transition: PPTSlideTransitionSchema.optional(),
})

export const PPTDeckSchema = z.object({
  id: z.string(),
  size: z.object({
    h: z.number(),
    w: z.number(),
  }),
  slides: z.array(PPTSlideSchema).min(1),
  title: z.string(),
})

export type PPTFill = z.infer<typeof PPTFillSchema>
export type PPTStroke = z.infer<typeof PPTStrokeSchema>
export type PPTStrokeDash = NonNullable<PPTStroke['dash']>
export type PPTGeometry = z.infer<typeof PPTGeometrySchema>
export type PPTRun = z.infer<typeof PPTRunSchema>
export type PPTParagraph = z.infer<typeof PPTParagraphSchema>
export type PPTTextBody = z.infer<typeof PPTTextBodySchema>
export type PPTTextStyle = z.infer<typeof PPTTextStyleSchema>
export type PPTTextAutoFit = z.infer<typeof PPTTextAutoFitSchema>
export type PPTSlideTransition = z.infer<typeof PPTSlideTransitionSchema>
export type PPTElementAccessibility = z.infer<typeof PPTElementAccessibilitySchema>
export type PPTElementAnimation = z.infer<typeof PPTElementAnimationSchema>
export type PPTElementShadow = z.infer<typeof PPTElementShadowSchema>
export type PPTElementHyperlink = z.infer<typeof PPTElementHyperlinkSchema>
export type PPTTextBox = z.infer<typeof PPTTextBoxSchema>
export type PPTShape = z.infer<typeof PPTShapeSchema>
export type PPTShapeKind = PPTShape['shape']
export type PPTImage = z.infer<typeof PPTImageSchema>
export type PPTImageAdjustments = NonNullable<PPTImage['adjustments']>
export type PPTImageClipShape = NonNullable<PPTImage['clipShape']>
export type PPTImageCrop = NonNullable<PPTImage['crop']>
export type PPTImageFit = NonNullable<PPTImage['fit']>
export type PPTLine = z.infer<typeof PPTLineSchema>
export type PPTLineConnection = z.infer<typeof PPTLineConnectionSchema>
export type PPTLineMarker = z.infer<typeof PPTLineMarkerSchema>
export type PPTLinePoint = z.infer<typeof PPTLinePointSchema>
export type PPTLineRoute = z.infer<typeof PPTLineRouteSchema>
export type PPTFreeform = z.infer<typeof PPTFreeformSchema>
export type PPTFreeformPathSegment =
  z.infer<typeof PPTFreeformPathSegmentSchema>
export type PPTTable = z.infer<typeof PPTTableSchema>
export type PPTTableCellBorders = z.infer<typeof PPTTableCellBordersSchema>
export type PPTTableCellStyle = z.infer<typeof PPTTableCellStyleSchema>
export type PPTTableCellTextStyle = z.infer<typeof PPTTableCellTextStyleSchema>
export type PPTComment = z.infer<typeof PPTCommentSchema>
export type PPTCommentThreadMessage = z.infer<typeof PPTCommentThreadMessageSchema>
export type PPTElement = z.infer<typeof PPTElementSchema>
export type PPTTextElement =
  | PPTTextBox
  | (PPTFreeform & { textBody: PPTTextBody })
  | (PPTShape & { textBody: PPTTextBody })
export type PPTSlide = z.infer<typeof PPTSlideSchema>
export type PPTDeck = z.infer<typeof PPTDeckSchema>

export function createPPTTextBody(text: string): PPTTextBody {
  return {
    paragraphs: text.split('\n').map((line) => ({
      runs: [{ text: line }],
    })),
  }
}

export function readPPTText(body: PPTTextBody | undefined) {
  if (!body) {
    return ''
  }

  return body.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n')
}

export function isPPTTextElement(
  element: PPTElement,
): element is PPTTextElement {
  return element.kind === 'textBox' ||
    (element.kind === 'freeform' && element.textBody !== undefined) ||
    (element.kind === 'shape' && element.textBody !== undefined)
}

export function replacePPTElementText(
  element: PPTElement,
  text: string,
): PPTElement {
  if (!isPPTTextElement(element)) {
    return element
  }

  return {
    ...element,
    textBody: copyPPTParagraphAttributes({
      source: element.textBody,
      target: createPPTTextBody(text),
    }),
  }
}

function copyPPTParagraphAttributes({
  source,
  target,
}: {
  source: PPTTextBody
  target: PPTTextBody
}): PPTTextBody {
  return {
    paragraphs: target.paragraphs.map((paragraph, index) => {
      const sourceParagraph = source.paragraphs[index]

      if (!sourceParagraph) {
        return paragraph
      }

      return {
        ...paragraph,
        ...(sourceParagraph.align ? { align: sourceParagraph.align } : {}),
        ...(sourceParagraph.bullet ? { bullet: sourceParagraph.bullet } : {}),
        ...(sourceParagraph.level === undefined ? {} : { level: sourceParagraph.level }),
        ...(sourceParagraph.lineHeight === undefined ? {} : { lineHeight: sourceParagraph.lineHeight }),
        ...(sourceParagraph.spacingAfter === undefined ? {} : { spacingAfter: sourceParagraph.spacingAfter }),
        ...(sourceParagraph.spacingBefore === undefined ? {} : { spacingBefore: sourceParagraph.spacingBefore }),
      }
    }),
  }
}

export function updatePPTElementGeometry(
  element: PPTElement,
  geometry: PPTGeometry,
): PPTElement {
  if (element.kind === 'line' || element.kind === 'freeform' || element.kind === 'table') {
    const scaleX = element.geometry.w === 0 ? 1 : geometry.w / element.geometry.w
    const scaleY = element.geometry.h === 0 ? 1 : geometry.h / element.geometry.h

    if (element.kind === 'freeform') {
      return {
        ...element,
        geometry,
        points: element.points.map((point) => ({
          x: point.x * scaleX,
          y: point.y * scaleY,
        })),
      }
    }

    if (element.kind === 'table') {
      return {
        ...element,
        ...(element.columnWidths
          ? { columnWidths: scalePPTTableTrackSizes(element.columnWidths, scaleX) }
          : {}),
        geometry,
        ...(element.rowHeights
          ? { rowHeights: scalePPTTableTrackSizes(element.rowHeights, scaleY) }
          : {}),
      }
    }

    return {
      ...element,
      end: {
        x: element.end.x * scaleX,
        y: element.end.y * scaleY,
      },
      geometry,
      start: {
        x: element.start.x * scaleX,
        y: element.start.y * scaleY,
      },
    }
  }

  return {
    ...element,
    geometry,
  }
}

function scalePPTTableTrackSizes(trackSizes: readonly number[], scale: number) {
  return trackSizes.map((size) => size * scale)
}

export function updatePPTDeckSlide(
  deck: PPTDeck,
  slideId: string,
  update: (slide: PPTSlide) => PPTSlide,
): PPTDeck {
  return {
    ...deck,
    slides: deck.slides.map((slide) =>
      slide.id === slideId ? update(slide) : slide,
    ),
  }
}

export function updatePPTDeckElement(
  deck: PPTDeck,
  slideId: string,
  elementId: string,
  update: (element: PPTElement) => PPTElement,
): PPTDeck {
  return updatePPTDeckSlide(deck, slideId, (slide) => ({
    ...slide,
    elements: slide.elements.map((element) =>
      element.id === elementId ? update(element) : element,
    ),
  }))
}

export function findPPTSlide(deck: PPTDeck, slideId: string) {
  return deck.slides.find((slide) => slide.id === slideId) ?? deck.slides[0]
}

export function findPPTElement(slide: PPTSlide, elementId: string | null) {
  if (!elementId) {
    return null
  }

  return slide.elements.find((element) => element.id === elementId) ?? null
}

export function createPPTElementId(slide: PPTSlide, prefix: string) {
  return createPPTCanvasSequentialIdFactory({
    existingIds: slide.elements.map((element) => element.id),
    formatId: ({ index, prefix }) => `${slide.id}-${prefix}-${index}`,
    startIndex: slide.elements.length + 1,
  })(prefix)
}
