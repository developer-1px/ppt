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
})

const PPTParagraphSchema = z.object({
  align: z.enum(['left', 'center', 'right']).optional(),
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
  id: z.string(),
  name: z.string(),
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

export const PPTElementSchema = z.discriminatedUnion('kind', [
  PPTTextBoxSchema,
  PPTShapeSchema,
  PPTImageSchema,
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
export type PPTImage = z.infer<typeof PPTImageSchema>
export type PPTElement = z.infer<typeof PPTElementSchema>
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
): element is PPTTextBox | PPTShape {
  return element.kind === 'textBox' ||
    (element.kind === 'shape' && element.textBody !== undefined)
}

export function replacePPTElementText(
  element: PPTElement,
  text: string,
): PPTElement {
  if (element.kind === 'image') {
    return element
  }

  return {
    ...element,
    textBody: createPPTTextBody(text),
  }
}

export function updatePPTElementGeometry(
  element: PPTElement,
  geometry: PPTGeometry,
): PPTElement {
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

