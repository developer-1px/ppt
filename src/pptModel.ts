import { z } from 'zod'

export const PPT_SLIDE_WIDTH = 1280
export const PPT_SLIDE_HEIGHT = 720

const PPTFillSchema = z.object({
  color: z.string(),
})

const PPTStrokeSchema = z.object({
  color: z.string(),
  width: z.number(),
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
  italic: z.boolean().optional(),
  size: z.number().optional(),
  text: z.string(),
  underline: z.boolean().optional(),
})

const PPTParagraphSchema = z.object({
  align: z.enum(['left', 'center', 'right']).optional(),
  bullet: z.enum(['bullet']).optional(),
  runs: z.array(PPTRunSchema),
})

export const PPTTextBodySchema = z.object({
  paragraphs: z.array(PPTParagraphSchema),
})

const PPTTextStyleSchema = z.object({
  color: z.string(),
  fontSize: z.number(),
  fontWeight: z.enum(['regular', 'semibold', 'bold']).optional(),
})

const PPTElementBaseSchema = z.object({
  geometry: PPTGeometrySchema,
  groupId: z.string().optional(),
  id: z.string(),
  locked: z.boolean().optional(),
  name: z.string(),
  visible: z.boolean().optional(),
})

const PPTTextBoxSchema = PPTElementBaseSchema.extend({
  kind: z.literal('textBox'),
  style: PPTTextStyleSchema,
  textBody: PPTTextBodySchema,
})

const PPTShapeSchema = PPTElementBaseSchema.extend({
  fill: PPTFillSchema,
  kind: z.literal('shape'),
  shape: z.enum(['rect', 'ellipse', 'diamond']),
  stroke: PPTStrokeSchema.optional(),
  style: PPTTextStyleSchema.optional(),
  textBody: PPTTextBodySchema.optional(),
})

const PPTImageSchema = PPTElementBaseSchema.extend({
  alt: z.string(),
  kind: z.literal('image'),
  src: z.string(),
})

const PPTLineMarkerSchema = z.enum(['none', 'arrow'])
const PPTLineRouteSchema = z.enum(['straight', 'elbow'])

const PPTLinePointSchema = z.object({
  x: z.number(),
  y: z.number(),
})

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

export const PPTElementSchema = z.discriminatedUnion('kind', [
  PPTTextBoxSchema,
  PPTShapeSchema,
  PPTImageSchema,
  PPTLineSchema,
])

export const PPTSlideSchema = z.object({
  background: PPTFillSchema.optional(),
  elements: z.array(PPTElementSchema),
  id: z.string(),
  name: z.string(),
  notes: z.string().optional(),
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
export type PPTGeometry = z.infer<typeof PPTGeometrySchema>
export type PPTRun = z.infer<typeof PPTRunSchema>
export type PPTParagraph = z.infer<typeof PPTParagraphSchema>
export type PPTTextBody = z.infer<typeof PPTTextBodySchema>
export type PPTTextStyle = z.infer<typeof PPTTextStyleSchema>
export type PPTTextBox = z.infer<typeof PPTTextBoxSchema>
export type PPTShape = z.infer<typeof PPTShapeSchema>
export type PPTShapeKind = PPTShape['shape']
export type PPTImage = z.infer<typeof PPTImageSchema>
export type PPTLine = z.infer<typeof PPTLineSchema>
export type PPTLineConnection = z.infer<typeof PPTLineConnectionSchema>
export type PPTLineMarker = z.infer<typeof PPTLineMarkerSchema>
export type PPTLinePoint = z.infer<typeof PPTLinePointSchema>
export type PPTLineRoute = z.infer<typeof PPTLineRouteSchema>
export type PPTElement = z.infer<typeof PPTElementSchema>
export type PPTTextElement = PPTTextBox | (PPTShape & { textBody: PPTTextBody })
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
      }
    }),
  }
}

export function updatePPTElementGeometry(
  element: PPTElement,
  geometry: PPTGeometry,
): PPTElement {
  if (element.kind === 'line') {
    const scaleX = element.geometry.w === 0 ? 1 : geometry.w / element.geometry.w
    const scaleY = element.geometry.h === 0 ? 1 : geometry.h / element.geometry.h

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
  const ids = new Set(slide.elements.map((element) => element.id))
  let next = slide.elements.length + 1
  let id = `${slide.id}-${prefix}-${next}`

  while (ids.has(id)) {
    next += 1
    id = `${slide.id}-${prefix}-${next}`
  }

  return id
}
