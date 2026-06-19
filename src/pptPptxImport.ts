import JSZip from 'jszip'
import {
  PPT_DEFAULT_THEME_ID,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  PPTDeckSchema,
  type PPTDeck,
  type PPTElement,
  type PPTFill,
  type PPTGeometry,
  type PPTImage,
  type PPTParagraph,
  type PPTRun,
  type PPTShapeKind,
  type PPTSlide,
  type PPTStroke,
  type PPTTextBody,
  type PPTTextStyle,
} from './pptModel'
import {
  PPTX_MIME_TYPE,
  PPTX_MODEL_CUSTOM_XML_CONTENT_TYPE,
  PPTX_MODEL_CUSTOM_XML_NAMESPACE,
  PPTX_MODEL_CUSTOM_XML_PATH,
} from './pptPptxExport'

export const PPTX_DECK_MODEL_IMPORT_FORMAT = 'pptx-custom-xml-ppt-deck' as const
export const PPTX_OPEN_XML_IMPORT_FORMAT = 'pptx-open-xml-ppt-deck' as const

type PPTDeckPPTXImportFormat =
  | typeof PPTX_DECK_MODEL_IMPORT_FORMAT
  | typeof PPTX_OPEN_XML_IMPORT_FORMAT

export type PPTDeckPPTXImportResult = {
  deck: PPTDeck
  format: PPTDeckPPTXImportFormat
  jsonLength: number
}

type PPTXRelationship = {
  target: string
  targetMode: string
  type: string
}
type PPTXRelationshipMap = Map<string, PPTXRelationship>

const PPTX_EMUS_PER_PIXEL = 9_525
const PPTX_TEXT_SIZE_UNITS_PER_POINT = 100
const PPTX_POINTS_PER_PIXEL = 0.75
const PPTX_DEFAULT_TEXT_COLOR = '#111827'
const PPTX_DEFAULT_TEXT_SIZE = 24
const PPTX_DEFAULT_FILL_COLOR = '#ffffff'
const PPTX_DEFAULT_STROKE_COLOR = '#111827'
const PPTX_SCHEME_COLORS: Record<string, string> = {
  accent1: '#2563eb',
  accent2: '#0ea5e9',
  accent3: '#22c55e',
  accent4: '#fb923c',
  accent5: '#7c2d12',
  accent6: '#dc2626',
  bg1: '#ffffff',
  bg2: '#f8fafc',
  dk1: '#111827',
  dk2: '#475569',
  lt1: '#ffffff',
  lt2: '#f8fafc',
  tx1: '#111827',
  tx2: '#475569',
}

export async function importPPTDeckFromPPTXBlob(
  blob: Blob,
): Promise<PPTDeckPPTXImportResult | null> {
  try {
    const zip = await JSZip.loadAsync(await blob.arrayBuffer())
    const embeddedModel = await importPPTDeckFromCustomXmlZip(zip)

    if (embeddedModel) {
      return embeddedModel
    }

    return await importPPTDeckFromOpenXmlZip(zip)
  } catch {
    return null
  }
}

async function importPPTDeckFromCustomXmlZip(
  zip: JSZip,
): Promise<PPTDeckPPTXImportResult | null> {
  const xml = await zip.file(PPTX_MODEL_CUSTOM_XML_PATH)?.async('string')

  if (!xml) {
    return null
  }

  const payload = getPPTDeckModelPayloadFromCustomXml(xml)

  if (!payload) {
    return null
  }

  try {
    const parsed = PPTDeckSchema.safeParse(JSON.parse(payload))

    if (parsed.success) {
      return {
        deck: parsed.data,
        format: PPTX_DECK_MODEL_IMPORT_FORMAT,
        jsonLength: payload.length,
      }
    }
  } catch {
    return null
  }

  return null
}

async function importPPTDeckFromOpenXmlZip(
  zip: JSZip,
): Promise<PPTDeckPPTXImportResult | null> {
  const slidePaths = await getPPTXOpenXmlSlidePaths(zip)

  if (slidePaths.length === 0) {
    return null
  }

  const size = await readPPTXOpenXmlDeckSize(zip)
  const title = await readPPTXOpenXmlDeckTitle(zip)
  const slides = await Promise.all(slidePaths.map((path, index) =>
    readPPTXOpenXmlSlide({
      index,
      path,
      zip,
    }),
  ))
  const parsed = PPTDeckSchema.safeParse({
    id: 'pptx-openxml-import',
    size,
    slides,
    title,
  })

  if (!parsed.success) {
    return null
  }

  return {
    deck: parsed.data,
    format: PPTX_OPEN_XML_IMPORT_FORMAT,
    jsonLength: 0,
  }
}

export function getPPTDeckPPTXFileFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTDeckPPTXFilesFromDataTransfer(dataTransfer)[0] ?? null
}

export function canImportPPTDeckPPTXFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTDeckPPTXFilesFromDataTransfer(dataTransfer).length > 0 ||
    hasPPTDeckPPTXItem(dataTransfer)
}

function getPPTDeckPPTXFilesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return Array.from(dataTransfer?.files ?? [])
    .filter(isPPTDeckPPTXFile)
}

function hasPPTDeckPPTXItem(dataTransfer: DataTransfer | null) {
  return Array.from(dataTransfer?.items ?? [])
    .some((item) => item.kind === 'file' && item.type === PPTX_MIME_TYPE)
}

function isPPTDeckPPTXFile(file: File) {
  return file.type === PPTX_MIME_TYPE ||
    file.name.toLowerCase().endsWith('.pptx')
}

async function getPPTXOpenXmlSlidePaths(zip: JSZip) {
  const pathsByFileName = getPPTXOpenXmlSlidePathsByFileName(zip)
  const orderedPaths = await readPPTXPresentationSlideOrder(zip)

  return orderedPaths.length > 0
    ? [
        ...orderedPaths,
        ...pathsByFileName.filter((path) => !orderedPaths.includes(path)),
      ]
    : pathsByFileName
}

function getPPTXOpenXmlSlidePathsByFileName(zip: JSZip) {
  return Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort(comparePPTXNumberedPaths)
}

async function readPPTXPresentationSlideOrder(zip: JSZip) {
  const presentationPath = 'ppt/presentation.xml'
  const xml = await zip.file(presentationPath)?.async('string')
  const doc = xml ? parsePPTXXmlDocument(xml) : null
  const relationships = await readPPTXRelationships(zip, presentationPath)

  if (!doc || relationships.size === 0) {
    return []
  }

  return getPPTXDescendantsByLocalName(doc, 'sldId')
    .map((slideId) =>
      slideId.getAttribute('r:id') ?? slideId.getAttribute('id'))
    .map((relationshipId) =>
      relationshipId ? relationships.get(relationshipId) : undefined)
    .filter((relationship): relationship is PPTXRelationship =>
      relationship !== undefined &&
      relationship.type.endsWith('/slide'))
    .map((relationship) =>
      resolvePPTXRelationshipTarget(presentationPath, relationship.target))
    .filter((path) => zip.file(path) !== null)
}

async function readPPTXOpenXmlDeckSize(zip: JSZip) {
  const presentationXml = await zip.file('ppt/presentation.xml')?.async('string')
  const doc = presentationXml ? parsePPTXXmlDocument(presentationXml) : null
  const slideSize = doc
    ? getFirstPPTXDescendantByLocalName(doc, 'sldSz')
    : null
  const width = toPPTXPositiveNumber(slideSize?.getAttribute('cx'))
  const height = toPPTXPositiveNumber(slideSize?.getAttribute('cy'))

  return {
    h: height === null ? PPT_SLIDE_HEIGHT : emuToPx(height),
    w: width === null ? PPT_SLIDE_WIDTH : emuToPx(width),
  }
}

async function readPPTXOpenXmlDeckTitle(zip: JSZip) {
  const coreXml = await zip.file('docProps/core.xml')?.async('string')
  const doc = coreXml ? parsePPTXXmlDocument(coreXml) : null
  const title = doc
    ? getFirstPPTXDescendantByLocalName(doc, 'title')?.textContent?.trim()
    : ''

  return title || 'Imported PPTX Deck'
}

async function readPPTXOpenXmlSlide({
  index,
  path,
  zip,
}: {
  index: number
  path: string
  zip: JSZip
}): Promise<PPTSlide> {
  const xml = await zip.file(path)?.async('string') ?? ''
  const doc = parsePPTXXmlDocument(xml)
  const cSld = doc ? getFirstPPTXDescendantByLocalName(doc, 'cSld') : null
  const spTree = cSld
    ? getFirstPPTXDescendantByLocalName(cSld, 'spTree')
    : null
  const relationships = await readPPTXSlideRelationships(zip, path)
  const notes = await readPPTXSlideNotes({
    relationships,
    slidePath: path,
    zip,
  })
  const elements: PPTElement[] = []
  let objectIndex = 1

  for (const child of getPPTXSlideObjectNodes(spTree, xml)) {
    const element = child.localName === 'sp'
      ? readPPTXShapeElement(child, index, objectIndex, relationships)
      : child.localName === 'pic'
        ? await readPPTXPictureElement({
            index,
            objectIndex,
            pic: child,
            relationships,
            slidePath: path,
            zip,
          })
        : null

    if (element) {
      elements.push(element)
      objectIndex += 1
    }
  }

  return {
    ...(readPPTXSlideBackground(cSld) ??
      readPPTXSlideBackgroundFromXml(xml) ?? {
        background: { color: PPTX_DEFAULT_FILL_COLOR },
      }),
    elements,
    id: `pptx-slide-${index + 1}`,
    name: readPPTXSlideName(cSld, index, xml),
    ...(notes ? { notes } : {}),
    themeId: PPT_DEFAULT_THEME_ID,
  }
}

function getPPTXSlideObjectNodes(spTree: Element | null, xml: string) {
  const treeNodes = spTree
    ? Array.from(spTree.children)
      .filter((child) => child.localName === 'sp' || child.localName === 'pic')
    : []

  if (treeNodes.length > 0) {
    return treeNodes
  }

  return [...xml.matchAll(/<p:(sp|pic)\b[\s\S]*?<\/p:\1>/g)]
    .map((match) => parsePPTXXmlElementFragment(match[0], match[1]))
    .filter((element): element is Element => element !== null)
}

function readPPTXSlideName(
  cSld: Element | null,
  index: number,
  xml: string,
) {
  const explicitName = cSld?.getAttribute('name')?.trim() ??
    xml.match(/<p:cSld\b[^>]*\bname="([^"]*)"/)?.[1]

  return explicitName
    ? unescapePPTXXmlAttribute(explicitName)
    : `Slide ${index + 1}`
}

function readPPTXSlideBackground(cSld: Element | null) {
  const bgPr = cSld
    ? getFirstPPTXDescendantByLocalName(cSld, 'bgPr')
    : null
  const fill = readPPTXSolidFill(bgPr)

  return fill ? { background: fill } : null
}

function readPPTXSlideBackgroundFromXml(xml: string) {
  const bgPrXml = xml.match(/<p:bgPr\b[\s\S]*?<\/p:bgPr>/)?.[0]
  const bgPr = bgPrXml ? parsePPTXXmlElementFragment(bgPrXml, 'bgPr') : null
  const fill = readPPTXSolidFill(bgPr)

  return fill ? { background: fill } : null
}

async function readPPTXSlideNotes({
  relationships,
  slidePath,
  zip,
}: {
  relationships: PPTXRelationshipMap
  slidePath: string
  zip: JSZip
}) {
  const notesRelationship = Array.from(relationships.values())
    .find((relationship) => relationship.type.endsWith('/notesSlide'))
  const notesPath = notesRelationship
    ? resolvePPTXRelationshipTarget(slidePath, notesRelationship.target)
    : null
  const xml = notesPath ? await zip.file(notesPath)?.async('string') : ''

  if (!xml) {
    return undefined
  }

  const doc = parsePPTXXmlDocument(xml)

  if (!doc) {
    return undefined
  }

  const bodyPlaceholder = getPPTXDescendantsByLocalName(doc, 'sp')
    .find((shape) =>
      getFirstPPTXDescendantByLocalName(shape, 'ph')
        ?.getAttribute('type') === 'body')
  const textBody = bodyPlaceholder
    ? getFirstPPTXDescendantByLocalName(bodyPlaceholder, 'txBody')
    : null
  const text = textBody
    ? readPPTXPlainTextBody(textBody)
    : readPPTXPlainTextBody(doc)

  return text.trim() || undefined
}

function readPPTXShapeElement(
  sp: Element,
  slideIndex: number,
  objectIndex: number,
  relationships: PPTXRelationshipMap,
): PPTElement | null {
  const spPr = getDirectPPTXChildByLocalName(sp, 'spPr')
  const txBody = getDirectPPTXChildByLocalName(sp, 'txBody')
  const geometry = readPPTXElementGeometry(spPr)
  const textBody = readPPTXTextBody(txBody)
  const fill = readPPTXSolidFill(spPr)
  const stroke = readPPTXStroke(spPr)
  const hasPaint = fill !== null || stroke !== undefined

  if (!geometry || (!textBody && !hasPaint)) {
    return null
  }

  const id = createPPTXImportedElementId(slideIndex, objectIndex)
  const name = readPPTXObjectName(sp, `Object ${objectIndex}`)
  const isTextBox = isPPTXTextBoxShape(sp) || (textBody !== null && !hasPaint)

  if (isTextBox) {
    return {
      geometry,
      ...(readPPTXElementHyperlink(sp, relationships) ?? {}),
      id,
      kind: 'textBox',
      name,
      style: readPPTXTextStyle(textBody, txBody),
      textAutoFit: 'resizeShapeToFitText',
      textBody: textBody ?? { paragraphs: [] },
    }
  }

  return {
    ...(readPPTXShapeCornerRadius(spPr) ?? {}),
    ...(readPPTXElementHyperlink(sp, relationships) ?? {}),
    ...(stroke ? { stroke } : {}),
    ...(textBody ? {
      style: readPPTXTextStyle(textBody, txBody),
      textAutoFit: 'resizeShapeToFitText' as const,
      textBody,
    } : {}),
    fill: fill ?? { color: PPTX_DEFAULT_FILL_COLOR },
    geometry,
    id,
    kind: 'shape',
    name,
    shape: readPPTXShapeKind(spPr),
  }
}

async function readPPTXPictureElement({
  index,
  objectIndex,
  pic,
  relationships,
  slidePath,
  zip,
}: {
  index: number
  objectIndex: number
  pic: Element
  relationships: PPTXRelationshipMap
  slidePath: string
  zip: JSZip
}): Promise<PPTImage | null> {
  const spPr = getDirectPPTXChildByLocalName(pic, 'spPr')
  const geometry = readPPTXElementGeometry(spPr)
  const blip = getFirstPPTXDescendantByLocalName(pic, 'blip')
  const relationshipId = blip?.getAttribute('r:embed') ??
    blip?.getAttribute('embed')
  const relationship = relationshipId ? relationships.get(relationshipId) : undefined
  const mediaPath = relationship
    ? resolvePPTXRelationshipTarget(slidePath, relationship.target)
    : null
  const media = mediaPath ? zip.file(mediaPath) : null

  if (!geometry || !mediaPath || !media) {
    return null
  }

  const base64 = await media.async('base64')
  const mimeType = getPPTXMediaMimeType(mediaPath)
  const name = readPPTXObjectName(pic, `Image ${objectIndex}`)
  const altText = readPPTXObjectDescription(pic)

  return {
    ...(altText ? { accessibility: { altText } } : {}),
    alt: altText || name,
    fit: 'contain',
    geometry,
    ...(readPPTXElementHyperlink(pic, relationships) ?? {}),
    id: createPPTXImportedElementId(index, objectIndex),
    kind: 'image',
    name,
    src: `data:${mimeType};base64,${base64}`,
  }
}

async function readPPTXSlideRelationships(zip: JSZip, slidePath: string) {
  return await readPPTXRelationships(zip, slidePath)
}

async function readPPTXRelationships(zip: JSZip, sourcePath: string) {
  const relationships = new Map<string, PPTXRelationship>()
  const relsPath = getPPTXRelationshipsPath(sourcePath)
  const xml = await zip.file(relsPath)?.async('string')
  const doc = xml ? parsePPTXXmlDocument(xml) : null

  if (!doc) {
    return relationships
  }

  for (const relationship of getPPTXDescendantsByLocalName(doc, 'Relationship')) {
    const id = relationship.getAttribute('Id')
    const target = relationship.getAttribute('Target')
    const targetMode = relationship.getAttribute('TargetMode') ?? ''
    const type = relationship.getAttribute('Type') ?? ''

    if (id && target) {
      relationships.set(id, {
        target,
        targetMode,
        type,
      })
    }
  }

  return relationships
}

function getPPTXRelationshipsPath(sourcePath: string) {
  const slashIndex = sourcePath.lastIndexOf('/')
  const directory = slashIndex >= 0 ? sourcePath.slice(0, slashIndex) : ''
  const fileName = slashIndex >= 0 ? sourcePath.slice(slashIndex + 1) : sourcePath

  return `${directory}/_rels/${fileName}.rels`
}

function resolvePPTXRelationshipTarget(basePath: string, target: string) {
  if (target.startsWith('/')) {
    return target.slice(1)
  }

  const baseDirectory = basePath.slice(0, basePath.lastIndexOf('/'))
  const parts = `${baseDirectory}/${target}`.split('/')
  const normalized: string[] = []

  for (const part of parts) {
    if (!part || part === '.') {
      continue
    }

    if (part === '..') {
      normalized.pop()
      continue
    }

    normalized.push(part)
  }

  return normalized.join('/')
}

function readPPTXElementGeometry(spPr: Element | null): PPTGeometry | null {
  const xfrm = spPr ? getDirectPPTXChildByLocalName(spPr, 'xfrm') : null
  const off = xfrm ? getDirectPPTXChildByLocalName(xfrm, 'off') : null
  const ext = xfrm ? getDirectPPTXChildByLocalName(xfrm, 'ext') : null
  const width = toPPTXPositiveNumber(ext?.getAttribute('cx'))
  const height = toPPTXPositiveNumber(ext?.getAttribute('cy'))

  if (!xfrm || width === null || height === null) {
    return null
  }

  const rotation = toPPTXNumber(xfrm.getAttribute('rot'))

  return {
    h: emuToPx(height),
    ...(rotation === null ? {} : { rotation: rotation / 60_000 }),
    w: emuToPx(width),
    x: emuToPx(toPPTXNumber(off?.getAttribute('x')) ?? 0),
    y: emuToPx(toPPTXNumber(off?.getAttribute('y')) ?? 0),
  }
}

function readPPTXTextBody(txBody: Element | null): PPTTextBody | null {
  if (!txBody) {
    return null
  }

  const paragraphs = getDirectPPTXChildrenByLocalName(txBody, 'p')
    .map(readPPTXParagraph)
  const hasText = paragraphs.some((paragraph) =>
    paragraph.runs.some((run) => run.text.length > 0))

  return hasText || paragraphs.length > 0 ? { paragraphs } : null
}

function readPPTXPlainTextBody(root: Document | Element) {
  return getPPTXDescendantsByLocalName(root, 'p')
    .map((paragraph) =>
      getPPTXDescendantsByLocalName(paragraph, 't')
        .map((text) => text.textContent ?? '')
        .join(''))
    .filter((text) => text.length > 0)
    .join('\n')
}

function readPPTXParagraph(paragraph: Element): PPTParagraph {
  const pPr = getDirectPPTXChildByLocalName(paragraph, 'pPr')
  const align = readPPTXParagraphAlign(pPr)
  const bullet = readPPTXParagraphBullet(pPr)
  const level = readPPTXParagraphLevel(pPr)
  const runs = Array.from(paragraph.children)
    .flatMap((child) => readPPTXTextRun(child))

  return {
    ...(align ? { align } : {}),
    ...(bullet ? { bullet } : {}),
    ...(level === undefined ? {} : { level }),
    runs: runs.length > 0 ? runs : [{ text: '' }],
  }
}

function readPPTXTextRun(node: Element): PPTRun[] {
  if (node.localName !== 'r' && node.localName !== 'fld') {
    return []
  }

  const text = getFirstPPTXDescendantByLocalName(node, 't')?.textContent ?? ''
  const rPr = getDirectPPTXChildByLocalName(node, 'rPr')
  const color = readPPTXSolidFill(rPr)?.color
  const size = toPPTXPositiveNumber(rPr?.getAttribute('sz'))

  return [{
    ...(rPr?.getAttribute('b') === '1' ? { bold: true } : {}),
    ...(color ? { color } : {}),
    ...(rPr?.getAttribute('i') === '1' ? { italic: true } : {}),
    ...(size === null ? {} : { size: textSizeToPx(size) }),
    ...(readPPTXUnderline(rPr) ? { underline: true } : {}),
    text,
  }]
}

function readPPTXTextStyle(
  textBody: PPTTextBody | null,
  txBody: Element | null,
): PPTTextStyle {
  const firstRun = textBody?.paragraphs
    .flatMap((paragraph) => paragraph.runs)
    .find((run) => run.text.trim().length > 0) ??
    textBody?.paragraphs[0]?.runs[0]
  const fontFamily = readPPTXFirstTypeface(txBody)

  return {
    color: firstRun?.color ?? PPTX_DEFAULT_TEXT_COLOR,
    ...(fontFamily ? { fontFamily } : {}),
    fontSize: firstRun?.size ?? PPTX_DEFAULT_TEXT_SIZE,
    ...(firstRun?.bold === true ? { fontWeight: 'bold' } : {}),
  }
}

function readPPTXParagraphAlign(
  pPr: Element | null,
): PPTParagraph['align'] | undefined {
  const align = pPr?.getAttribute('algn')

  if (align === 'ctr') {
    return 'center'
  }

  if (align === 'r') {
    return 'right'
  }

  if (align === 'just') {
    return 'justify'
  }

  return align === 'l' ? 'left' : undefined
}

function readPPTXParagraphBullet(
  pPr: Element | null,
): PPTParagraph['bullet'] | undefined {
  if (!pPr) {
    return undefined
  }

  if (getDirectPPTXChildByLocalName(pPr, 'buAutoNum')) {
    return 'numbered'
  }

  return getDirectPPTXChildByLocalName(pPr, 'buChar') ? 'bullet' : undefined
}

function readPPTXParagraphLevel(pPr: Element | null) {
  const level = toPPTXNumber(pPr?.getAttribute('lvl'))

  return level === null ? undefined : Math.max(0, level)
}

function readPPTXUnderline(rPr: Element | null) {
  const underline = rPr?.getAttribute('u')

  return underline !== null && underline !== undefined && underline !== 'none'
}

function readPPTXTypeface(rPr: Element | null) {
  const latin = rPr ? getDirectPPTXChildByLocalName(rPr, 'latin') : null
  const typeface = latin?.getAttribute('typeface')?.trim()

  return typeface || undefined
}

function readPPTXFirstTypeface(txBody: Element | null) {
  if (!txBody) {
    return undefined
  }

  for (const runProperties of getPPTXDescendantsByLocalName(txBody, 'rPr')) {
    const typeface = readPPTXTypeface(runProperties)

    if (typeface) {
      return typeface
    }
  }

  return undefined
}

function isPPTXTextBoxShape(sp: Element) {
  return getFirstPPTXDescendantByLocalName(sp, 'cNvSpPr')
    ?.getAttribute('txBox') === '1'
}

function readPPTXShapeKind(spPr: Element | null): PPTShapeKind {
  const preset = getFirstPPTXDescendantByLocalName(spPr, 'prstGeom')
    ?.getAttribute('prst')

  if (preset === 'ellipse') {
    return 'ellipse'
  }

  if (preset === 'diamond') {
    return 'diamond'
  }

  return 'rect'
}

function readPPTXShapeCornerRadius(spPr: Element | null) {
  const preset = getFirstPPTXDescendantByLocalName(spPr, 'prstGeom')
    ?.getAttribute('prst')

  return preset === 'roundRect' ? { cornerRadius: 24 } : null
}

function readPPTXSolidFill(container: Element | null): PPTFill | null {
  if (!container || getDirectPPTXChildByLocalName(container, 'noFill')) {
    return null
  }

  const solidFill = getDirectPPTXChildByLocalName(container, 'solidFill')

  if (!solidFill) {
    return null
  }

  const color = readPPTXColor(solidFill)

  if (!color) {
    return null
  }

  const opacity = readPPTXAlphaOpacity(solidFill)

  return {
    color,
    ...(opacity === null ? {} : { opacity }),
  }
}

function readPPTXStroke(spPr: Element | null): PPTStroke | undefined {
  const line = spPr ? getDirectPPTXChildByLocalName(spPr, 'ln') : null

  if (!line || getDirectPPTXChildByLocalName(line, 'noFill')) {
    return undefined
  }

  const dash = readPPTXStrokeDash(line)

  return {
    ...(dash ? { dash } : {}),
    color: readPPTXSolidFill(line)?.color ?? PPTX_DEFAULT_STROKE_COLOR,
    width: Math.max(1, emuToPx(toPPTXPositiveNumber(line.getAttribute('w')) ?? PPTX_EMUS_PER_PIXEL)),
  }
}

function readPPTXStrokeDash(line: Element): PPTStroke['dash'] | undefined {
  const value = getDirectPPTXChildByLocalName(line, 'prstDash')
    ?.getAttribute('val')

  if (value === 'dot' || value === 'sysDot') {
    return 'dot'
  }

  return value === 'dash' || value === 'lgDash' ? 'dash' : undefined
}

function readPPTXColor(solidFill: Element) {
  const srgb = getDirectPPTXChildByLocalName(solidFill, 'srgbClr')
    ?.getAttribute('val')
  const scheme = getDirectPPTXChildByLocalName(solidFill, 'schemeClr')
    ?.getAttribute('val')

  if (srgb && /^[\da-f]{6}$/i.test(srgb)) {
    return `#${srgb.toLowerCase()}`
  }

  return scheme ? PPTX_SCHEME_COLORS[scheme] : undefined
}

function readPPTXAlphaOpacity(solidFill: Element) {
  const alpha = getFirstPPTXDescendantByLocalName(solidFill, 'alpha')
  const value = toPPTXPositiveNumber(alpha?.getAttribute('val'))

  return value === null ? null : Math.max(0, Math.min(1, value / 100_000))
}

function readPPTXObjectName(element: Element, fallback: string) {
  return getFirstPPTXDescendantByLocalName(element, 'cNvPr')
    ?.getAttribute('name')
    ?.trim() || fallback
}

function readPPTXObjectDescription(element: Element) {
  return getFirstPPTXDescendantByLocalName(element, 'cNvPr')
    ?.getAttribute('descr')
    ?.trim() ?? ''
}

function readPPTXElementHyperlink(
  element: Element,
  relationships: PPTXRelationshipMap,
) {
  const relationshipId = getFirstPPTXDescendantByLocalName(element, 'hlinkClick')
    ?.getAttribute('r:id')
  const relationship = relationshipId
    ? relationships.get(relationshipId)
    : undefined
  const url = relationship?.targetMode === 'External'
    ? relationship.target
    : undefined

  return url ? { hyperlink: { url } } : null
}

function createPPTXImportedElementId(slideIndex: number, objectIndex: number) {
  return `pptx-slide-${slideIndex + 1}-object-${objectIndex}`
}

function getPPTXMediaMimeType(path: string) {
  const extension = path.split('.').at(-1)?.toLowerCase()

  if (extension === 'svg') {
    return 'image/svg+xml'
  }

  if (extension === 'jpg' || extension === 'jpeg') {
    return 'image/jpeg'
  }

  if (extension === 'gif') {
    return 'image/gif'
  }

  if (extension === 'webp') {
    return 'image/webp'
  }

  return 'image/png'
}

function parsePPTXXmlDocument(xml: string) {
  if (typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(xml, 'application/xml')

  return getPPTXDescendantsByLocalName(doc, 'parsererror').length > 0
    ? null
    : doc
}

function parsePPTXXmlElementFragment(xml: string, localName: string) {
  const doc = parsePPTXXmlDocument([
    '<pptxFragment ',
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ',
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
    xml,
    '</pptxFragment>',
  ].join(''))

  return doc ? getFirstPPTXDescendantByLocalName(doc, localName) : null
}

function getFirstPPTXDescendantByLocalName(
  root: Document | Element | null,
  localName: string,
) {
  return root ? getPPTXDescendantsByLocalName(root, localName)[0] ?? null : null
}

function getPPTXDescendantsByLocalName(
  root: Document | Element,
  localName: string,
) {
  return Array.from(root.getElementsByTagName('*'))
    .filter((element) => element.localName === localName)
}

function getDirectPPTXChildByLocalName(
  element: Element | null,
  localName: string,
) {
  return getDirectPPTXChildrenByLocalName(element, localName)[0] ?? null
}

function getDirectPPTXChildrenByLocalName(
  element: Element | null,
  localName: string,
) {
  return element
    ? Array.from(element.children)
      .filter((child) => child.localName === localName)
    : []
}

function comparePPTXNumberedPaths(left: string, right: string) {
  return getPPTXPathNumber(left) - getPPTXPathNumber(right)
}

function getPPTXPathNumber(path: string) {
  return Number(path.match(/(\d+)\.xml$/)?.[1] ?? 0)
}

function toPPTXNumber(value: string | null | undefined) {
  if (value === null || value === undefined || value.trim() === '') {
    return null
  }

  const parsed = Number(value)

  return Number.isFinite(parsed) ? parsed : null
}

function toPPTXPositiveNumber(value: string | null | undefined) {
  const parsed = toPPTXNumber(value)

  return parsed === null || parsed < 0 ? null : parsed
}

function emuToPx(value: number) {
  return Math.round(value / PPTX_EMUS_PER_PIXEL)
}

function textSizeToPx(value: number) {
  return Math.round(
    value / PPTX_TEXT_SIZE_UNITS_PER_POINT / PPTX_POINTS_PER_PIXEL,
  )
}

function getPPTDeckModelPayloadFromCustomXml(xml: string) {
  const match = xml.match(/<pptDeck\b([^>]*)>([\s\S]*?)<\/pptDeck>/)

  if (!match) {
    return null
  }

  const attributes = match[1]

  if (!attributes.includes(`xmlns="${PPTX_MODEL_CUSTOM_XML_NAMESPACE}"`) ||
    !attributes.includes(`contentType="${PPTX_MODEL_CUSTOM_XML_CONTENT_TYPE}"`)) {
    return null
  }

  return unescapePPTXXmlText(match[2]).trim()
}

function unescapePPTXXmlText(value: string) {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}

function unescapePPTXXmlAttribute(value: string) {
  return unescapePPTXXmlText(value)
    .replaceAll('&quot;', '"')
    .replaceAll('&apos;', "'")
}
