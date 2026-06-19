import PptxGenJS from 'pptxgenjs'
import {
  type PPTComment,
  type PPTDeck,
  type PPTElement,
  type PPTElementShadow,
  type PPTFill,
  type PPTFreeform,
  type PPTGeometry,
  type PPTImage,
  type PPTLine,
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTSlide,
  type PPTStroke,
  type PPTTable,
  type PPTTextBody,
  type PPTTextAutoFit,
  type PPTTextStyle,
} from './pptModel'

export const PPTX_MIME_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

const PPTX_PIXELS_PER_INCH = 96
const PPTX_POINTS_PER_PIXEL = 0.75
const PPTX_DEFAULT_TEXT_COLOR = '#111827'
const PPTX_DEFAULT_FONT_FACE = 'Inter'
const PPTX_DEFAULT_FONT_SIZE = 24
const PPTX_DEFAULT_SHAPE_TEXT_INSET = Object.freeze({
  bottom: 18,
  left: 18,
  right: 18,
  top: 18,
} as const)
const PPTX_DEFAULT_TEXT_BOX_INSET = Object.freeze({
  bottom: 0,
  left: 0,
  right: 0,
  top: 0,
} as const)
const PPTX_DEFAULT_SHAPE_CORNER_RADIUS = 24

type PPTXPptx = InstanceType<typeof PptxGenJS>
type PPTXSlide = ReturnType<PPTXPptx['addSlide']>
type PPTXPosition = {
  h: number
  rotate?: number
  w: number
  x: number
  y: number
}
type PPTXTextInset = {
  bottom: number
  left: number
  right: number
  top: number
}

export function createPPTDeckPPTX(deck: PPTDeck) {
  const pptx = new PptxGenJS()
  const layoutName = getPPTXLayoutName(deck)

  pptx.author = 'Interactive OS'
  pptx.company = 'Interactive OS'
  pptx.subject = 'PPT deck export'
  pptx.title = deck.title
  pptx.theme = {
    bodyFontFace: PPTX_DEFAULT_FONT_FACE,
    headFontFace: PPTX_DEFAULT_FONT_FACE,
  }
  pptx.defineLayout({
    height: deck.size.h / PPTX_PIXELS_PER_INCH,
    name: layoutName,
    width: deck.size.w / PPTX_PIXELS_PER_INCH,
  })
  pptx.layout = layoutName

  for (const deckSlide of deck.slides) {
    addPPTXSlide({
      pptxSlide: pptx.addSlide(),
      slide: deckSlide,
    })
  }

  return pptx
}

export async function exportPPTDeckPPTXBlob(deck: PPTDeck) {
  const output = await createPPTDeckPPTX(deck).write({
    compression: true,
    outputType: 'blob',
  })

  if (output instanceof Blob) {
    return output.type === PPTX_MIME_TYPE
      ? output
      : new Blob([output], {
          type: PPTX_MIME_TYPE,
        })
  }

  const blobPart = typeof output === 'string' || output instanceof ArrayBuffer
    ? output
    : new Uint8Array(output).buffer

  return new Blob([blobPart], {
    type: PPTX_MIME_TYPE,
  })
}

export function getPPTDeckPPTXFilename(deck: Pick<PPTDeck, 'title'>) {
  const slug = deck.title
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')

  return `${slug || 'ppt-deck'}.pptx`
}

function addPPTXSlide({
  pptxSlide,
  slide,
}: {
  pptxSlide: PPTXSlide
  slide: PPTSlide
}) {
  pptxSlide.background = {
    color: toPPTXColor(slide.background?.color ?? '#ffffff', 'FFFFFF'),
  }

  for (const element of slide.elements) {
    if (element.visible === false) {
      continue
    }

    addPPTXElement({
      element,
      pptxSlide,
    })
  }

  const notes = slide.notes?.trim()

  if (notes) {
    pptxSlide.addNotes(notes)
  }
}

function addPPTXElement({
  element,
  pptxSlide,
}: {
  element: PPTElement
  pptxSlide: PPTXSlide
}) {
  if (element.kind === 'textBox') {
    addPPTXTextBox({
      element,
      pptxSlide,
    })
    return
  }

  if (element.kind === 'shape') {
    addPPTXShape({
      element,
      pptxSlide,
    })
    return
  }

  if (element.kind === 'image') {
    addPPTXImage({
      element,
      pptxSlide,
    })
    return
  }

  if (element.kind === 'line') {
    addPPTXLine({
      element,
      pptxSlide,
    })
    return
  }

  if (element.kind === 'freeform') {
    addPPTXFreeform({
      element,
      pptxSlide,
    })
    return
  }

  if (element.kind === 'table') {
    addPPTXTable({
      element,
      pptxSlide,
    })
    return
  }

  addPPTXComment({
    element,
    pptxSlide,
  })
}

function addPPTXTextBox({
  element,
  pptxSlide,
}: {
  element: Extract<PPTElement, { kind: 'textBox' }>
  pptxSlide: PPTXSlide
}) {
  pptxSlide.addText(createPPTXTextRuns({
    body: element.textBody,
    style: element.style,
  }), {
    ...createPPTXElementTextOptions({
      element,
      inset: getPPTXTextInset(element.style, PPTX_DEFAULT_TEXT_BOX_INSET),
    }),
    isTextBox: true,
  })
}

function addPPTXShape({
  element,
  pptxSlide,
}: {
  element: PPTShape
  pptxSlide: PPTXSlide
}) {
  const shapeName = getPPTXShapeName(element)
  const shapeOptions = {
    ...createPPTXElementPosition(element.geometry),
    fill: createPPTXFill(element.fill),
    flipH: element.flipH === true,
    flipV: element.flipV === true,
    hyperlink: createPPTXHyperlink(element.hyperlink),
    line: createPPTXLineProps(element.stroke),
    objectName: element.name,
    rectRadius: getPPTXRectRadius(element),
    shadow: createPPTXShadow(element.shadow),
  }

  if (element.textBody) {
    pptxSlide.addText(createPPTXTextRuns({
      body: element.textBody,
      style: element.style,
    }), {
      ...shapeOptions,
      ...createPPTXTextStyleOptions(element.style),
      fit: getPPTXTextFit(element.textAutoFit),
      margin: createPPTXMargin(getPPTXTextInset(
        element.style,
        PPTX_DEFAULT_SHAPE_TEXT_INSET,
      )),
      shape: shapeName,
      valign: element.style?.verticalAlign ?? 'top',
    })
    return
  }

  pptxSlide.addShape(shapeName, shapeOptions)
}

function addPPTXImage({
  element,
  pptxSlide,
}: {
  element: PPTImage
  pptxSlide: PPTXSlide
}) {
  const src = element.src.trim()
  const source = src.startsWith('data:')
    ? { data: src }
    : { path: src }
  const position = createPPTXElementPosition(element.geometry)

  pptxSlide.addImage({
    ...source,
    ...position,
    altText: element.accessibility?.altText ?? element.alt,
    flipH: element.flipH === true,
    flipV: element.flipV === true,
    hyperlink: createPPTXHyperlink(element.hyperlink),
    objectName: element.name,
    rotate: element.geometry.rotation,
    shadow: createPPTXShadow(element.shadow),
    sizing: {
      h: position.h,
      type: element.fit === 'contain' ? 'contain' : 'cover',
      w: position.w,
    },
    transparency: toPPTXTransparency(element.opacity ?? 1),
  })
}

function addPPTXLine({
  element,
  pptxSlide,
}: {
  element: PPTLine
  pptxSlide: PPTXSlide
}) {
  const start = getPPTXLineWorldPoint(element, element.start)
  const end = getPPTXLineWorldPoint(element, element.end)

  if (element.route === 'elbow') {
    addPPTXElbowRoute({
      element,
      pptxSlide,
      start,
      end,
    })
    return
  }

  pptxSlide.addShape('line', {
    h: pxToIn(end.y - start.y),
    line: createPPTXLineProps(element.stroke, {
      endMarker: element.endMarker,
      startMarker: element.startMarker,
    }),
    objectName: element.name,
    rotate: element.geometry.rotation,
    shadow: createPPTXShadow(element.shadow),
    w: pxToIn(end.x - start.x),
    x: pxToIn(start.x),
    y: pxToIn(start.y),
  })

}

function addPPTXElbowRoute({
  element,
  end,
  pptxSlide,
  start,
}: {
  element: PPTLine
  end: { x: number, y: number }
  pptxSlide: PPTXSlide
  start: { x: number, y: number }
}) {
  const bendX = element.geometry.x +
    element.geometry.w * (element.routeBend ?? 0.5)
  const segments = [
    [start, { x: bendX, y: start.y }],
    [{ x: bendX, y: start.y }, { x: bendX, y: end.y }],
    [{ x: bendX, y: end.y }, end],
  ] as const

  segments.forEach(([segmentStart, segmentEnd], index) => {
    pptxSlide.addShape('line', {
      h: pxToIn(segmentEnd.y - segmentStart.y),
      line: createPPTXLineProps(element.stroke, {
        endMarker: index === segments.length - 1 ? element.endMarker : undefined,
        startMarker: index === 0 ? element.startMarker : undefined,
      }),
      objectName: `${element.name} route`,
      w: pxToIn(segmentEnd.x - segmentStart.x),
      x: pxToIn(segmentStart.x),
      y: pxToIn(segmentStart.y),
    })
  })
}

function addPPTXFreeform({
  element,
  pptxSlide,
}: {
  element: PPTFreeform
  pptxSlide: PPTXSlide
}) {
  for (let index = 1; index < element.points.length; index += 1) {
    const start = element.points[index - 1]
    const end = element.points[index]

    pptxSlide.addShape('line', {
      h: pxToIn(end.y - start.y),
      line: createPPTXLineProps(element.stroke),
      objectName: `${element.name} segment ${index}`,
      w: pxToIn(end.x - start.x),
      x: pxToIn(element.geometry.x + start.x),
      y: pxToIn(element.geometry.y + start.y),
    })
  }

}

function addPPTXTable({
  element,
  pptxSlide,
}: {
  element: PPTTable
  pptxSlide: PPTXSlide
}) {
  const columnCount = Math.max(
    1,
    ...element.rows.map((row) => row.length),
  )
  const position = createPPTXElementPosition(element.geometry)
  const colW = Array.from(
    { length: columnCount },
    () => position.w / columnCount,
  )
  const rowH = element.rows.length > 0
    ? Array.from({ length: element.rows.length }, () => position.h / element.rows.length)
    : undefined

  pptxSlide.addTable(element.rows.map((row, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) => ({
      options: {
        bold: rowIndex === 0,
        border: { color: 'DBE3EF', pt: 0.75 },
        color: '111827',
        fill: { color: rowIndex === 0 ? 'EFF6FF' : 'FFFFFF' },
        fontFace: PPTX_DEFAULT_FONT_FACE,
        fontSize: 13.5,
        margin: 0.08,
      },
      text: row[columnIndex] ?? '',
    }))),
  {
    ...position,
    border: { color: 'DBE3EF', pt: 0.75 },
    colW,
    objectName: element.name,
    rowH,
  })
}

function addPPTXComment({
  element,
  pptxSlide,
}: {
  element: PPTComment
  pptxSlide: PPTXSlide
}) {
  const body = [
    element.body,
    ...(element.thread ?? []).map((message) =>
      `${message.authorName}: ${message.body}`),
  ].filter(Boolean).join('\n')

  pptxSlide.addText(body, {
    ...createPPTXElementPosition(element.geometry),
    color: '78350F',
    fill: { color: element.resolved ? 'FEF3C7' : 'FFFBEB' },
    fontFace: PPTX_DEFAULT_FONT_FACE,
    fontSize: 12.75,
    line: { color: 'D97706', width: 0.75 },
    margin: [6, 8, 6, 8],
    objectName: element.name,
    shape: 'roundRect',
  })
}

function createPPTXElementTextOptions({
  element,
  inset,
}: {
  element: Extract<PPTElement, { kind: 'textBox' }>
  inset: PPTXTextInset
}) {
  return {
    ...createPPTXElementPosition(element.geometry),
    ...createPPTXTextStyleOptions(element.style),
    fit: getPPTXTextFit(element.textAutoFit),
    hyperlink: createPPTXHyperlink(element.hyperlink),
    margin: createPPTXMargin(inset),
    objectName: element.name,
    shadow: createPPTXShadow(element.shadow),
    valign: element.style.verticalAlign ?? 'top',
  }
}

function createPPTXTextRuns({
  body,
  style,
}: {
  body: PPTTextBody
  style?: PPTTextStyle
}): PptxGenJS.TextProps[] {
  const runs: PptxGenJS.TextProps[] = []

  body.paragraphs.forEach((paragraph, paragraphIndex) => {
    const paragraphRuns = paragraph.runs.length > 0
      ? paragraph.runs
      : [{ text: '' }]

    paragraphRuns.forEach((run, runIndex) => {
      runs.push({
        options: {
          ...createPPTXTextStyleOptions(style),
          ...createPPTXParagraphOptions(paragraph),
          ...createPPTXTextRunOptions(run),
          breakLine: paragraphIndex > 0 && runIndex === 0,
        },
        text: run.text,
      })
    })
  })

  return runs.length > 0
    ? runs
    : [{ text: '' }]
}

function createPPTXParagraphOptions(
  paragraph: PPTParagraph,
): Partial<PptxGenJS.TextPropsOptions> {
  const listLevel = Math.max(0, paragraph.level ?? 0)
  const bullet: PptxGenJS.TextPropsOptions['bullet'] = paragraph.bullet
    ? {
        indent: 14 + listLevel * 18,
        type: paragraph.bullet === 'numbered' ? 'number' : 'bullet',
      }
    : undefined

  return {
    align: paragraph.align,
    bullet,
    indentLevel: listLevel,
    lineSpacingMultiple: paragraph.lineHeight,
    paraSpaceAfter: paragraph.spacingAfter === undefined
      ? undefined
      : pxToPt(paragraph.spacingAfter),
    paraSpaceBefore: paragraph.spacingBefore === undefined
      ? undefined
      : pxToPt(paragraph.spacingBefore),
  }
}

function createPPTXTextRunOptions(
  run: PPTRun,
): Partial<PptxGenJS.TextPropsOptions> {
  const strike: PptxGenJS.TextPropsOptions['strike'] = run.strikethrough === true
    ? 'sngStrike'
    : undefined

  return {
    bold: run.bold === true,
    color: run.color ? toPPTXColor(run.color, '111827') : undefined,
    fontSize: run.size === undefined ? undefined : pxToPt(run.size),
    highlight: run.highlight ? toPPTXColor(run.highlight, 'FEF08A') : undefined,
    italic: run.italic === true,
    strike,
    underline: run.underline === true ? { style: 'sng' } : undefined,
  }
}

function createPPTXTextStyleOptions(
  style: PPTTextStyle | undefined,
): Partial<PptxGenJS.TextPropsOptions> {
  return {
    bold: style?.fontWeight === 'bold',
    color: toPPTXColor(style?.color ?? PPTX_DEFAULT_TEXT_COLOR, '111827'),
    fontFace: style?.fontFamily ?? PPTX_DEFAULT_FONT_FACE,
    fontSize: pxToPt(style?.fontSize ?? PPTX_DEFAULT_FONT_SIZE),
  }
}

function createPPTXElementPosition(
  geometry: PPTGeometry,
): PPTXPosition {
  return {
    h: pxToIn(geometry.h),
    rotate: geometry.rotation,
    w: pxToIn(geometry.w),
    x: pxToIn(geometry.x),
    y: pxToIn(geometry.y),
  }
}

function createPPTXFill(fill: PPTFill) {
  return {
    color: toPPTXColor(fill.color, 'FFFFFF'),
    transparency: toPPTXTransparency(fill.opacity ?? 1),
  }
}

function createPPTXLineProps(
  stroke: PPTStroke | undefined,
  markers: {
    endMarker?: PPTLine['endMarker']
    startMarker?: PPTLine['startMarker']
  } = {},
): PptxGenJS.ShapeLineProps {
  if (!stroke) {
    return {
      color: 'FFFFFF',
      transparency: 100,
      width: 0,
    }
  }

  const dashType: PptxGenJS.ShapeLineProps['dashType'] = stroke.dash === 'dot'
    ? 'sysDot'
    : stroke.dash === 'dash'
      ? 'dash'
      : 'solid'

  return {
    beginArrowType: markers.startMarker === 'arrow' ? 'arrow' : undefined,
    color: toPPTXColor(stroke.color, '000000'),
    dashType,
    endArrowType: markers.endMarker === 'arrow' ? 'arrow' : undefined,
    width: Math.max(0.25, pxToPt(stroke.width)),
  }
}

function createPPTXHyperlink(
  hyperlink: PPTElement['hyperlink'] | undefined,
) {
  return hyperlink?.url
    ? { url: hyperlink.url }
    : undefined
}

function createPPTXShadow(shadow: PPTElementShadow | undefined) {
  if (!shadow) {
    return undefined
  }

  return {
    angle: shadow.angle,
    blur: pxToPt(shadow.blur),
    color: toPPTXColor(shadow.color, '000000'),
    offset: pxToPt(shadow.distance),
    opacity: clamp(shadow.opacity, 0, 1),
    type: 'outer',
  } satisfies PptxGenJS.ShadowProps
}

function createPPTXMargin(inset: PPTXTextInset): [number, number, number, number] {
  return [
    pxToPt(inset.top),
    pxToPt(inset.right),
    pxToPt(inset.bottom),
    pxToPt(inset.left),
  ]
}

function getPPTXTextInset(
  style: PPTTextStyle | undefined,
  fallback: PPTXTextInset,
) {
  return style?.textInset ?? fallback
}

function getPPTXTextFit(
  textAutoFit: PPTTextAutoFit | undefined,
): PptxGenJS.TextPropsOptions['fit'] {
  return textAutoFit === 'resizeShapeToFitText'
    ? 'resize'
    : 'none'
}

function getPPTXShapeName(element: PPTShape): PptxGenJS.SHAPE_NAME {
  if (element.shape === 'ellipse') {
    return 'ellipse'
  }

  if (element.shape === 'diamond') {
    return 'diamond'
  }

  return getPPTXShapeCornerRadius(element) > 0
    ? 'roundRect'
    : 'rect'
}

function getPPTXRectRadius(element: PPTShape) {
  if (element.shape !== 'rect') {
    return undefined
  }

  const radius = getPPTXShapeCornerRadius(element)

  if (radius <= 0) {
    return undefined
  }

  return clamp(radius / Math.max(1, Math.min(element.geometry.w, element.geometry.h)), 0, 1)
}

function getPPTXShapeCornerRadius(element: PPTShape) {
  return element.cornerRadius ?? PPTX_DEFAULT_SHAPE_CORNER_RADIUS
}

function getPPTXLineWorldPoint(
  element: PPTLine,
  point: PPTLine['start'],
) {
  return {
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  }
}

function getPPTXLayoutName(deck: PPTDeck) {
  return `PPT_MODEL_${Math.round(deck.size.w)}x${Math.round(deck.size.h)}`
}

function toPPTXColor(color: string, fallback: string) {
  const trimmed = color.trim()
  const hex = trimmed.startsWith('#') ? trimmed.slice(1) : trimmed

  if (/^[0-9a-f]{3}$/i.test(hex)) {
    return hex
      .split('')
      .map((char) => `${char}${char}`)
      .join('')
      .toUpperCase()
  }

  return /^[0-9a-f]{6}$/i.test(hex)
    ? hex.toUpperCase()
    : fallback
}

function toPPTXTransparency(opacity: number) {
  return Math.round((1 - clamp(opacity, 0, 1)) * 100)
}

function pxToIn(value: number) {
  return value / PPTX_PIXELS_PER_INCH
}

function pxToPt(value: number) {
  return value * PPTX_POINTS_PER_PIXEL
}

function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}
