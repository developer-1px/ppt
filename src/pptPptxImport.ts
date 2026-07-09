import JSZip from 'jszip'
import {
  PPT_DEFAULT_THEME_ID,
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  PPTDeckSchema,
  type PPTDeck,
  type PPTElement,
  type PPTElementAnimation,
  type PPTElementShadow,
  type PPTFill,
  type PPTGeometry,
  type PPTImage,
  type PPTLine,
  type PPTLineConnection,
  type PPTParagraph,
  type PPTRun,
  type PPTShapeKind,
  type PPTSlide,
  type PPTSlideTransition,
  type PPTStroke,
  type PPTTable,
  type PPTTableCellBorders,
  type PPTTableCellStyle,
  type PPTTableCellTextStyle,
  type PPTTextAutoFit,
  type PPTTextBody,
  type PPTTextStyle,
} from './pptModel'
import { getPPTTableColumnCount } from './pptTableLayout'
import {
  PPTX_MIME_TYPE,
  PPTX_MODEL_CUSTOM_XML_CONTENT_TYPE,
  PPTX_MODEL_CUSTOM_XML_NAMESPACE,
  PPTX_MODEL_CUSTOM_XML_PATH,
} from './pptPptxExport'

export const PPTX_DECK_MODEL_IMPORT_FORMAT = 'pptx-custom-xml-ppt-deck' as const
export const PPTX_OPEN_XML_IMPORT_FORMAT = 'pptx-open-xml-ppt-deck' as const
const PPTX_RELATIONSHIP_ATTRIBUTE_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PPTX_ROUND_RECT_DEFAULT_ADJUST = 16_667
const PPTX_ROUND_RECT_MAX_ADJUST = 50_000

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
type PPTXImportedAnimation = {
  animation: PPTElementAnimation
  objectName?: string
  objectId: string
}
type PPTXGroupTransform = {
  childOffsetX: number
  childOffsetY: number
  offsetX: number
  offsetY: number
  scaleX: number
  scaleY: number
}
type PPTXSlideObjectNode = {
  element: Element
  groupId?: string
  transform: PPTXGroupTransform
}
type PPTXLineConnectionRef = {
  anchor: PPTLineConnection['anchor']
  objectId: string
}
type PPTXLineConnectionRefs = {
  end?: PPTXLineConnectionRef
  start?: PPTXLineConnectionRef
}
type PPTXThemeColorMap = Readonly<Record<string, string>>
type PPTXThemeFontMap = Readonly<Record<string, string>>

const PPTX_EMUS_PER_PIXEL = 9_525
const PPTX_TEXT_SIZE_UNITS_PER_POINT = 100
const PPTX_POINTS_PER_PIXEL = 0.75
const PPTX_DEFAULT_TEXT_COLOR = '#111827'
const PPTX_DEFAULT_TEXT_SIZE = 24
const PPTX_DEFAULT_FILL_COLOR = '#ffffff'
const PPTX_DEFAULT_STROKE_COLOR = '#111827'
const PPTX_IDENTITY_GROUP_TRANSFORM: PPTXGroupTransform = {
  childOffsetX: 0,
  childOffsetY: 0,
  offsetX: 0,
  offsetY: 0,
  scaleX: 1,
  scaleY: 1,
}
const PPTX_LOCK_ATTRIBUTE_NAMES = [
  'noAdjustHandles',
  'noEditPoints',
  'noMove',
  'noResize',
  'noRot',
  'noSelect',
  'noTextEdit',
] as const
const PPTX_LOCK_TAG_NAMES = [
  'cxnSpLocks',
  'graphicFrameLocks',
  'grpSpLocks',
  'picLocks',
  'spLocks',
] as const
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
const PPTX_PRESET_COLORS: Record<string, string> = {
  black: '#000000',
  blue: '#0000ff',
  cyan: '#00ffff',
  dkBlue: '#00008b',
  dkCyan: '#008b8b',
  dkGray: '#a9a9a9',
  dkGreen: '#006400',
  dkMagenta: '#8b008b',
  dkRed: '#8b0000',
  dkYellow: '#808000',
  gold: '#ffd700',
  green: '#008000',
  ltBlue: '#add8e6',
  ltCyan: '#e0ffff',
  ltGray: '#d3d3d3',
  ltGreen: '#90ee90',
  ltMagenta: '#ff77ff',
  ltYellow: '#ffffe0',
  magenta: '#ff00ff',
  orange: '#ffa500',
  purple: '#800080',
  red: '#ff0000',
  white: '#ffffff',
  yellow: '#ffff00',
}
const PPTX_SYSTEM_COLORS: Record<string, string> = {
  '3dDkShadow': '#696969',
  '3dLight': '#e3e3e3',
  activeBorder: '#b4b4b4',
  activeCaption: '#99b4d1',
  appWorkspace: '#ababab',
  background: '#000000',
  btnFace: '#f0f0f0',
  btnHighlight: '#ffffff',
  btnShadow: '#a0a0a0',
  btnText: '#000000',
  captionText: '#000000',
  gradientActiveCaption: '#b9d1ea',
  gradientInactiveCaption: '#d7e4f2',
  grayText: '#6d6d6d',
  highlight: '#0078d7',
  highlightText: '#ffffff',
  hotLight: '#0066cc',
  inactiveBorder: '#f4f7fc',
  inactiveCaption: '#bfcddb',
  inactiveCaptionText: '#000000',
  infoBk: '#ffffe1',
  infoText: '#000000',
  menu: '#f0f0f0',
  menuBar: '#f0f0f0',
  menuHighlight: '#3399ff',
  menuText: '#000000',
  scrollBar: '#c8c8c8',
  window: '#ffffff',
  windowFrame: '#646464',
  windowText: '#000000',
}
const PPTX_THEME_SCHEME_ALIASES: Record<string, string> = {
  bg1: 'lt1',
  bg2: 'lt2',
  tx1: 'dk1',
  tx2: 'dk2',
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
  const themePath = await readPPTXOpenXmlThemePath(zip)
  const themeColors = await readPPTXOpenXmlThemeColors(zip, themePath)
  const themeFonts = await readPPTXOpenXmlThemeFonts(zip, themePath)
  const title = await readPPTXOpenXmlDeckTitle(zip)
  const slides = await Promise.all(slidePaths.map((path, index) =>
    readPPTXOpenXmlSlide({
      index,
      path,
      themeColors,
      themeFonts,
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

export function getPPTDeckPPTXFileFromList(files: FileList | null) {
  return getPPTDeckPPTXFilesFromList(files)[0] ?? null
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
  return getPPTDeckPPTXFilesFromList(dataTransfer?.files ?? null)
}

function getPPTDeckPPTXFilesFromList(files: FileList | null) {
  return Array.from(files ?? []).filter(isPPTDeckPPTXFile)
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

async function readPPTXOpenXmlThemeColors(
  zip: JSZip,
  themePath: string | null,
): Promise<PPTXThemeColorMap> {
  const xml = themePath ? await zip.file(themePath)?.async('string') : null
  const doc = xml ? parsePPTXXmlDocument(xml) : null
  const colorScheme = doc
    ? getFirstPPTXDescendantByLocalName(doc, 'clrScheme')
    : null

  if (!colorScheme) {
    return PPTX_SCHEME_COLORS
  }

  const colors = Array.from(colorScheme.children)
    .reduce<Record<string, string>>((next, colorNode) => {
      const color = readPPTXThemeColorNode(colorNode)

      return color
        ? { ...next, [colorNode.localName]: color }
        : next
    }, {})

  return resolvePPTXThemeSchemeColors(colors)
}

async function readPPTXOpenXmlThemeFonts(
  zip: JSZip,
  themePath: string | null,
): Promise<PPTXThemeFontMap> {
  const xml = themePath ? await zip.file(themePath)?.async('string') : null
  const doc = xml ? parsePPTXXmlDocument(xml) : null
  const fontScheme = doc
    ? getFirstPPTXDescendantByLocalName(doc, 'fontScheme')
    : null
  const majorFont = getDirectPPTXChildByLocalName(fontScheme, 'majorFont')
  const minorFont = getDirectPPTXChildByLocalName(fontScheme, 'minorFont')

  return {
    ...readPPTXThemeFontGroup(majorFont, '+mj'),
    ...readPPTXThemeFontGroup(minorFont, '+mn'),
  }
}

function readPPTXThemeFontGroup(
  fontGroup: Element | null,
  prefix: '+mj' | '+mn',
): Record<string, string> {
  const latin = readPPTXThemeTypeface(fontGroup, 'latin')
  const eastAsian = readPPTXThemeTypeface(fontGroup, 'ea') ?? latin
  const complexScript = readPPTXThemeTypeface(fontGroup, 'cs') ?? latin

  return {
    ...(latin ? { [`${prefix}-lt`]: latin } : {}),
    ...(eastAsian ? { [`${prefix}-ea`]: eastAsian } : {}),
    ...(complexScript ? { [`${prefix}-cs`]: complexScript } : {}),
  }
}

function readPPTXThemeTypeface(
  fontGroup: Element | null,
  localName: 'cs' | 'ea' | 'latin',
) {
  const typeface = getDirectPPTXChildByLocalName(fontGroup, localName)
    ?.getAttribute('typeface')
    ?.trim()

  return typeface && !typeface.startsWith('+') ? typeface : undefined
}

async function readPPTXOpenXmlThemePath(zip: JSZip) {
  const presentationPath = 'ppt/presentation.xml'
  const relationships = await readPPTXRelationships(zip, presentationPath)
  const themeRelationship = Array.from(relationships.values())
    .find((relationship) =>
      relationship.targetMode !== 'External' &&
      relationship.type.endsWith('/theme'))
  const relatedThemePath = themeRelationship
    ? resolvePPTXRelationshipTarget(presentationPath, themeRelationship.target)
    : null

  if (relatedThemePath && zip.file(relatedThemePath)) {
    return relatedThemePath
  }

  return Object.keys(zip.files)
    .filter((path) => /^ppt\/theme\/theme\d+\.xml$/.test(path))
    .sort(comparePPTXNumberedPaths)[0] ?? null
}

function readPPTXThemeColorNode(colorNode: Element) {
  const srgbColor = getDirectPPTXChildByLocalName(colorNode, 'srgbClr')
  const presetColor = getDirectPPTXChildByLocalName(colorNode, 'prstClr')
  const systemColor = getDirectPPTXChildByLocalName(colorNode, 'sysClr')
  const color = [
    {
      color: readPPTXHexColor(srgbColor?.getAttribute('val')),
      element: srgbColor,
    },
    {
      color: readPPTXPresetColor(presetColor?.getAttribute('val')),
      element: presetColor,
    },
    {
      color: readPPTXHexColor(systemColor?.getAttribute('lastClr')) ??
        readPPTXSystemColor(systemColor?.getAttribute('val')),
      element: systemColor,
    },
  ].find((candidate) => candidate.color && candidate.element)

  return color?.color && color.element
    ? applyPPTXColorModifiers(color.color, color.element)
    : color?.color
}

function resolvePPTXThemeSchemeColors(
  themeColors: Readonly<Record<string, string>>,
): PPTXThemeColorMap {
  const colors = { ...PPTX_SCHEME_COLORS, ...themeColors }

  for (const [alias, source] of Object.entries(PPTX_THEME_SCHEME_ALIASES)) {
    colors[alias] = themeColors[source] ?? themeColors[alias] ?? colors[alias]
  }

  return colors
}

async function readPPTXOpenXmlSlide({
  index,
  path,
  themeColors,
  themeFonts,
  zip,
}: {
  index: number
  path: string
  themeColors: PPTXThemeColorMap
  themeFonts: PPTXThemeFontMap
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
  const transition = readPPTXSlideTransition(doc, xml)
  const elements: PPTElement[] = []
  const elementIdByPptxObjectId = new Map<string, string>()
  const lineConnectionRefsByElementId = new Map<string, PPTXLineConnectionRefs>()
  let objectIndex = 1

  for (const objectNode of getPPTXSlideObjectNodes(spTree, xml, index)) {
    const child = objectNode.element
    let element: PPTElement | null = null

    if (child.localName === 'sp') {
      element = isPPTXLineShape(child)
        ? readPPTXLineElement(child, index, objectIndex, relationships, themeColors)
        : readPPTXShapeElement(child, index, objectIndex, relationships, themeColors, themeFonts)
    } else if (child.localName === 'cxnSp') {
      element = readPPTXLineElement(child, index, objectIndex, relationships, themeColors)
    } else if (child.localName === 'pic') {
      element = await readPPTXPictureElement({
        index,
        objectIndex,
        pic: child,
        relationships,
        slidePath: path,
        themeColors,
        zip,
      })
    } else if (child.localName === 'graphicFrame') {
      element = readPPTXTableElement(child, index, objectIndex, relationships, themeColors)
    }

    if (element) {
      const transformedElement = applyPPTXGroupObjectNode(element, objectNode)

      elements.push(transformedElement)
      if (transformedElement.kind === 'line') {
        const lineConnectionRefs = readPPTXLineConnectionRefs(child)

        if (lineConnectionRefs) {
          lineConnectionRefsByElementId.set(transformedElement.id, lineConnectionRefs)
        }
      }
      for (const pptxObjectId of readPPTXObjectIds(child)) {
        elementIdByPptxObjectId.set(pptxObjectId, transformedElement.id)
      }
      objectIndex += 1
    }
  }
  const connectedElements = applyPPTXLineConnections({
    elementIdByPptxObjectId,
    elements,
    lineConnectionRefsByElementId,
  })
  const animatedElements = applyPPTXElementAnimations({
    elementIdByPptxObjectId,
    elements: connectedElements,
    importedAnimations: readPPTXSlideAnimations(doc, xml),
  })

  return {
    ...(readPPTXSlideBackground(cSld, themeColors) ??
      readPPTXSlideBackgroundFromXml(xml, themeColors) ?? {
        background: { color: PPTX_DEFAULT_FILL_COLOR },
      }),
    elements: animatedElements,
    id: `pptx-slide-${index + 1}`,
    name: readPPTXSlideName(cSld, index, xml),
    ...(notes ? { notes } : {}),
    themeId: PPT_DEFAULT_THEME_ID,
    ...(transition ? { transition } : {}),
  }
}

function applyPPTXLineConnections({
  elementIdByPptxObjectId,
  elements,
  lineConnectionRefsByElementId,
}: {
  elementIdByPptxObjectId: ReadonlyMap<string, string>
  elements: readonly PPTElement[]
  lineConnectionRefsByElementId: ReadonlyMap<string, PPTXLineConnectionRefs>
}) {
  return elements.map((element) => {
    if (element.kind !== 'line') {
      return element
    }

    const refs = lineConnectionRefsByElementId.get(element.id)

    if (!refs) {
      return element
    }

    const startConnection = resolvePPTXLineConnection(
      refs.start,
      elementIdByPptxObjectId,
    )
    const endConnection = resolvePPTXLineConnection(
      refs.end,
      elementIdByPptxObjectId,
    )

    return {
      ...element,
      ...(startConnection ? { startConnection } : {}),
      ...(endConnection ? { endConnection } : {}),
    }
  })
}

function resolvePPTXLineConnection(
  connection: PPTXLineConnectionRef | undefined,
  elementIdByPptxObjectId: ReadonlyMap<string, string>,
): PPTLineConnection | undefined {
  if (!connection) {
    return undefined
  }

  const elementId = elementIdByPptxObjectId.get(connection.objectId)

  return elementId ? { anchor: connection.anchor, elementId } : undefined
}

function applyPPTXElementAnimations({
  elementIdByPptxObjectId,
  elements,
  importedAnimations,
}: {
  elementIdByPptxObjectId: ReadonlyMap<string, string>
  elements: readonly PPTElement[]
  importedAnimations: readonly PPTXImportedAnimation[]
}) {
  const animationByElementId = new Map<string, PPTElementAnimation>()
  const elementIdByObjectName = new Map(
    elements.map((element) => [element.name, element.id]),
  )

  importedAnimations.forEach((imported, index) => {
    const elementId = elementIdByPptxObjectId.get(imported.objectId) ??
      (imported.objectName
        ? elementIdByObjectName.get(imported.objectName)
        : undefined)

    if (elementId && !animationByElementId.has(elementId)) {
      animationByElementId.set(elementId, {
        ...imported.animation,
        order: index + 1,
      })
    }
  })

  return elements.map((element) => {
    const animation = animationByElementId.get(element.id)

    return animation ? { ...element, animation } : element
  })
}

function readPPTXObjectIds(element: Element) {
  return getPPTXDescendantsByLocalName(element, 'cNvPr')
    .map((nonVisualProperties) =>
      nonVisualProperties.getAttribute('id')?.trim() ?? '')
    .filter((id) => id.length > 0)
}

function applyPPTXGroupObjectNode(
  element: PPTElement,
  objectNode: PPTXSlideObjectNode,
): PPTElement {
  const withGroup = objectNode.groupId
    ? { ...element, groupId: objectNode.groupId }
    : element

  return isPPTXIdentityGroupTransform(objectNode.transform)
    ? withGroup
    : transformPPTXElement(withGroup, objectNode.transform)
}

function getPPTXSlideObjectNodes(
  spTree: Element | null,
  xml: string,
  slideIndex: number,
) {
  const treeNodes = spTree
    ? getPPTXSlideObjectNodesFromContainer({
        container: spTree,
        groupId: undefined,
        slideIndex,
        transform: PPTX_IDENTITY_GROUP_TRANSFORM,
      })
    : []

  if (treeNodes.length > 0) {
    return treeNodes
  }

  return [...xml.matchAll(/<p:(sp|cxnSp|pic|graphicFrame)\b[\s\S]*?<\/p:\1>/g)]
    .map((match) => parsePPTXXmlElementFragment(match[0], match[1]))
    .filter((element): element is Element => element !== null)
    .map((element) => ({
      element,
      transform: PPTX_IDENTITY_GROUP_TRANSFORM,
    }))
}

function getPPTXSlideObjectNodesFromContainer({
  container,
  groupId,
  slideIndex,
  transform,
}: {
  container: Element
  groupId: string | undefined
  slideIndex: number
  transform: PPTXGroupTransform
}): PPTXSlideObjectNode[] {
  return Array.from(container.children).flatMap((child) => {
    if (isPPTXSlideObjectNode(child)) {
      return [{
        element: child,
        ...(groupId ? { groupId } : {}),
        transform,
      }]
    }

    if (child.localName !== 'grpSp') {
      return []
    }

    const nextTransform = composePPTXGroupTransforms(
      transform,
      readPPTXGroupTransform(child),
    )
    const nextGroupId = groupId ??
      readPPTXGroupId(child, slideIndex)

    return getPPTXSlideObjectNodesFromContainer({
      container: child,
      groupId: nextGroupId,
      slideIndex,
      transform: nextTransform,
    })
  })
}

function isPPTXSlideObjectNode(element: Element) {
  return element.localName === 'sp' ||
    element.localName === 'cxnSp' ||
    element.localName === 'pic' ||
    element.localName === 'graphicFrame'
}

function readPPTXGroupId(group: Element, slideIndex: number) {
  const id = getFirstPPTXDescendantByLocalName(group, 'cNvPr')
    ?.getAttribute('id')
    ?.trim()
  const name = readPPTXObjectName(group, 'group')
    .replace(/[^\w.-]+/g, '-')
    .replace(/^-+|-+$/g, '')

  return `pptx-slide-${slideIndex + 1}-group-${id || name || 'object'}`
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

function readPPTXSlideBackground(
  cSld: Element | null,
  themeColors: PPTXThemeColorMap,
) {
  const background = cSld
    ? getDirectPPTXChildByLocalName(cSld, 'bg')
    : null
  const bgPr = getDirectPPTXChildByLocalName(background, 'bgPr') ??
    (cSld ? getFirstPPTXDescendantByLocalName(cSld, 'bgPr') : null)
  const bgRef = getDirectPPTXChildByLocalName(background, 'bgRef') ??
    (cSld ? getFirstPPTXDescendantByLocalName(cSld, 'bgRef') : null)
  const fill = readPPTXFill(bgPr, themeColors) ??
    readPPTXBackgroundRefFill(bgRef, themeColors)

  return fill ? { background: fill } : null
}

function readPPTXSlideBackgroundFromXml(
  xml: string,
  themeColors: PPTXThemeColorMap,
) {
  const bgPrXml = xml.match(/<p:bgPr\b[\s\S]*?<\/p:bgPr>/)?.[0]
  const bgRefXml = xml.match(/<p:bgRef\b[\s\S]*?<\/p:bgRef>/)?.[0]
  const bgPr = bgPrXml ? parsePPTXXmlElementFragment(bgPrXml, 'bgPr') : null
  const bgRef = bgRefXml ? parsePPTXXmlElementFragment(bgRefXml, 'bgRef') : null
  const fill = readPPTXFill(bgPr, themeColors) ??
    readPPTXBackgroundRefFill(bgRef, themeColors)

  return fill ? { background: fill } : null
}

function readPPTXBackgroundRefFill(
  bgRef: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  return readPPTXColorFill(bgRef, themeColors)
}

function readPPTXSlideTransition(
  doc: Document | null,
  xml: string,
): PPTSlideTransition | null {
  const transition = doc
    ? getFirstPPTXDescendantByLocalName(doc, 'transition')
    : null

  if (!transition) {
    return readPPTXSlideTransitionFromXml(xml)
  }

  return readPPTXSlideTransitionElement(transition)
}

function readPPTXSlideTransitionElement(
  transition: Element,
): PPTSlideTransition {
  return {
    advanceAfterMs: toPPTXPositiveNumber(transition.getAttribute('advTm')),
    advanceOnClick: transition.getAttribute('advClick') !== '0',
    durationMs: readPPTXSlideTransitionDuration(transition),
    type: readPPTXSlideTransitionType(transition),
  }
}

function readPPTXSlideTransitionFromXml(xml: string): PPTSlideTransition | null {
  const match = xml.match(/<p:transition\b([^>]*)>([\s\S]*?)<\/p:transition>|<p:transition\b([^>]*)\/>/)

  if (!match) {
    return null
  }

  const attributes = match[1] ?? match[3] ?? ''
  const body = match[2] ?? ''

  return {
    advanceAfterMs: toPPTXPositiveNumber(readPPTXXmlAttribute(attributes, 'advTm')),
    advanceOnClick: readPPTXXmlAttribute(attributes, 'advClick') !== '0',
    durationMs: toPPTXPositiveNumber(
      readPPTXXmlAttribute(attributes, 'dur') ??
        readPPTXXmlAttribute(attributes, 'p14:dur'),
    ) ?? readPPTXSlideTransitionSpeedDurationFromValue(
      readPPTXXmlAttribute(attributes, 'spd'),
    ),
    type: body.includes('<p:push')
      ? 'push'
      : body.includes('<p:fade')
        ? 'fade'
        : 'none',
  }
}

function readPPTXSlideTransitionDuration(transition: Element) {
  return toPPTXPositiveNumber(
    transition.getAttribute('p14:dur') ??
      getPPTXAttributeByLocalName(transition, 'dur'),
  ) ?? readPPTXSlideTransitionSpeedDurationFromValue(
    transition.getAttribute('spd'),
  )
}

function readPPTXSlideTransitionSpeedDurationFromValue(
  speed: string | null,
) {
  if (speed === 'fast') {
    return 500
  }

  if (speed === 'slow') {
    return 1500
  }

  return 650
}

function readPPTXSlideTransitionType(
  transition: Element,
): PPTSlideTransition['type'] {
  if (getDirectPPTXChildByLocalName(transition, 'push')) {
    return 'push'
  }

  return getDirectPPTXChildByLocalName(transition, 'fade') ? 'fade' : 'none'
}

function readPPTXSlideAnimations(
  doc: Document | null,
  xml: string,
): PPTXImportedAnimation[] {
  const objectNameById = readPPTXObjectNameByIdFromXml(xml)
  const fromDom = doc
    ? Array.from(doc.getElementsByTagName('*'))
      .filter((element) =>
        element.localName === 'animEffect' ||
        element.localName === 'animMotion')
      .map(readPPTXAnimationEffect)
      .filter((animation): animation is PPTXImportedAnimation => animation !== null)
    : []
  const fromXml = readPPTXSlideAnimationsFromXml(xml)
  const seen = new Set<string>()

  return [...fromDom, ...fromXml]
    .map((animation) => ({
      ...animation,
      ...(objectNameById.get(animation.objectId)
        ? { objectName: objectNameById.get(animation.objectId) }
        : {}),
    }))
    .filter((animation) => {
      const key = [
        animation.objectId,
        animation.animation.delayMs,
        animation.animation.durationMs,
        animation.animation.trigger,
        animation.animation.type,
      ].join(':')

      if (seen.has(key)) {
        return false
      }

      seen.add(key)
      return true
    })
}

function readPPTXAnimationEffect(effect: Element): PPTXImportedAnimation | null {
  const objectId = getFirstPPTXDescendantByLocalName(effect, 'spTgt')
    ?.getAttribute('spid')
    ?.trim()
  const behavior = getDirectPPTXChildByLocalName(effect, 'cBhvr') ??
    getFirstPPTXDescendantByLocalName(effect, 'cBhvr')
  const behaviorTiming = getDirectPPTXChildByLocalName(behavior, 'cTn')
  const containerTiming = findPPTXAncestorByLocalName(effect, 'cTn')

  if (!objectId || !behaviorTiming || !containerTiming) {
    return null
  }

  const trigger: PPTElementAnimation['trigger'] =
    containerTiming.getAttribute('nodeType') === 'withEffect'
      ? 'withPrevious'
      : 'onClick'
  const behaviorDelayMs = readPPTXAnimationDelayMs(behaviorTiming) ?? 0
  const containerDelayMs = readPPTXAnimationDelayMs(containerTiming) ?? 0
  const type = readPPTXAnimationType(effect)

  if (!type) {
    return null
  }

  return {
    animation: {
      delayMs: trigger === 'withPrevious' ? containerDelayMs : behaviorDelayMs,
      durationMs: toPPTXPositiveNumber(behaviorTiming.getAttribute('dur')) ?? 500,
      order: 1,
      trigger,
      type,
    },
    objectId,
  }
}

function readPPTXAnimationType(
  effect: Element,
): PPTElementAnimation['type'] | null {
  if (effect.localName === 'animMotion') {
    return 'flyIn'
  }

  const filter = effect.getAttribute('filter')?.toLowerCase() ?? ''
  const transition = effect.getAttribute('transition')

  return effect.localName === 'animEffect' &&
    transition === 'in' &&
    filter.includes('fade')
    ? 'fadeIn'
    : null
}

function readPPTXAnimationDelayMs(timing: Element) {
  const stCondLst = getDirectPPTXChildByLocalName(timing, 'stCondLst')
  const condition = getDirectPPTXChildByLocalName(stCondLst, 'cond')

  return toPPTXPositiveNumber(condition?.getAttribute('delay'))
}

function readPPTXSlideAnimationsFromXml(xml: string): PPTXImportedAnimation[] {
  const animations: PPTXImportedAnimation[] = []

  for (const match of xml.matchAll(/<p:par>([\s\S]*?)<\/p:par>/g)) {
    const block = match[1]

    if (!block.includes('<p:animEffect') && !block.includes('<p:animMotion')) {
      continue
    }

    const animation = readPPTXAnimationEffectFromXml(block)

    if (animation) {
      animations.push(animation)
    }
  }

  return animations
}

function readPPTXAnimationEffectFromXml(
  block: string,
): PPTXImportedAnimation | null {
  const targetAttributes = block.match(/<p:spTgt\b([^>]*)\/>/)?.[1] ?? ''
  const objectId = readPPTXXmlAttribute(targetAttributes, 'spid')?.trim()
  const effectMatch = block.match(/<p:(animEffect|animMotion)\b([^>]*)>([\s\S]*?)<\/p:\1>/)
  const containerAttributes = block.match(/<p:cTn\b([^>]*)>/)?.[1] ?? ''
  const behaviorMatch = effectMatch?.[3]
    .match(/<p:cBhvr>[\s\S]*?<p:cTn\b([^>]*)>([\s\S]*?)<\/p:cTn>/)
  const behaviorAttributes = behaviorMatch?.[1] ?? ''
  const behaviorBody = behaviorMatch?.[2] ?? ''

  if (!objectId || !effectMatch || !behaviorMatch) {
    return null
  }

  const type = readPPTXAnimationTypeFromXml(
    effectMatch[1],
    effectMatch[2],
  )

  if (!type) {
    return null
  }

  const trigger: PPTElementAnimation['trigger'] =
    readPPTXXmlAttribute(containerAttributes, 'nodeType') === 'withEffect'
      ? 'withPrevious'
      : 'onClick'
  const behaviorDelayMs = readPPTXAnimationDelayMsFromXml(behaviorBody) ?? 0
  const containerDelayMs = readPPTXAnimationDelayMsFromXml(block) ?? 0

  return {
    animation: {
      delayMs: trigger === 'withPrevious' ? containerDelayMs : behaviorDelayMs,
      durationMs: toPPTXPositiveNumber(
        readPPTXXmlAttribute(behaviorAttributes, 'dur'),
      ) ?? 500,
      order: 1,
      trigger,
      type,
    },
    objectId,
  }
}

function readPPTXAnimationTypeFromXml(
  tagName: string,
  attributes: string,
): PPTElementAnimation['type'] | null {
  if (tagName === 'animMotion') {
    return 'flyIn'
  }

  const filter = readPPTXXmlAttribute(attributes, 'filter')?.toLowerCase() ?? ''
  const transition = readPPTXXmlAttribute(attributes, 'transition')

  return tagName === 'animEffect' &&
    transition === 'in' &&
    filter.includes('fade')
    ? 'fadeIn'
    : null
}

function readPPTXAnimationDelayMsFromXml(xml: string) {
  const conditionAttributes = xml.match(/<p:cond\b([^>]*)\/>/)?.[1] ?? ''

  return toPPTXPositiveNumber(
    readPPTXXmlAttribute(conditionAttributes, 'delay'),
  )
}

function readPPTXObjectNameByIdFromXml(xml: string) {
  const objectNameById = new Map<string, string>()

  for (const match of xml.matchAll(/<p:cNvPr\b([^>]*)>/g)) {
    const id = readPPTXXmlAttribute(match[1], 'id')?.trim()
    const name = readPPTXXmlAttribute(match[1], 'name')?.trim()

    if (id && name) {
      objectNameById.set(id, unescapePPTXXmlAttribute(name))
    }
  }

  return objectNameById
}

function findPPTXAncestorByLocalName(
  element: Element,
  localName: string,
) {
  let current = getPPTXParentElement(element)

  while (current) {
    if (current.localName === localName) {
      return current
    }

    current = getPPTXParentElement(current)
  }

  return null
}

function getPPTXParentElement(element: Element) {
  const parent = element.parentElement ?? element.parentNode

  return parent instanceof Element ? parent : null
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

function readPPTXLineElement(
  element: Element,
  slideIndex: number,
  objectIndex: number,
  relationships: PPTXRelationshipMap,
  themeColors: PPTXThemeColorMap,
): PPTElement | null {
  const spPr = getDirectPPTXChildByLocalName(element, 'spPr')
  const line = getDirectPPTXChildByLocalName(spPr, 'ln')
  const stroke = readPPTXStroke(spPr, themeColors)
  const opacity = readPPTXLineOpacity(line)
  const shadow = readPPTXElementShadow(spPr, themeColors)
  const lineGeometry = readPPTXLineGeometry(spPr)

  if (!stroke || !lineGeometry) {
    return null
  }

  return {
    ...(readPPTXElementAccessibility(element) ?? {}),
    end: lineGeometry.end,
    endMarker: readPPTXLineMarker(line, 'tailEnd'),
    geometry: lineGeometry.geometry,
    ...(readPPTXElementHyperlink(element, relationships) ?? {}),
    id: createPPTXImportedElementId(slideIndex, objectIndex),
    kind: 'line',
    ...(readPPTXElementLocked(element) ? { locked: true } : {}),
    ...(readPPTXElementVisibility(element) ?? {}),
    name: readPPTXObjectName(element, `Line ${objectIndex}`),
    ...(opacity === null ? {} : { opacity }),
    route: readPPTXLineRoute(spPr),
    ...(shadow ? { shadow } : {}),
    start: lineGeometry.start,
    startMarker: readPPTXLineMarker(line, 'headEnd'),
    stroke,
  }
}

function isPPTXLineShape(sp: Element) {
  const spPr = getDirectPPTXChildByLocalName(sp, 'spPr')
  const preset = getFirstPPTXDescendantByLocalName(spPr, 'prstGeom')
    ?.getAttribute('prst')

  return preset === 'line' || preset === 'straightConnector1'
}

function readPPTXLineRoute(spPr: Element | null): PPTLine['route'] {
  const preset = getFirstPPTXDescendantByLocalName(spPr, 'prstGeom')
    ?.getAttribute('prst')

  return preset?.startsWith('bentConnector') ? 'elbow' : 'straight'
}

function readPPTXLineConnectionRefs(
  element: Element,
): PPTXLineConnectionRefs | null {
  const nonVisualConnection = getFirstPPTXDescendantByLocalName(element, 'cNvCxnSpPr')
  const start = readPPTXLineConnectionRef(
    getDirectPPTXChildByLocalName(nonVisualConnection, 'stCxn'),
  )
  const end = readPPTXLineConnectionRef(
    getDirectPPTXChildByLocalName(nonVisualConnection, 'endCxn'),
  )

  return start || end
    ? {
        ...(start ? { start } : {}),
        ...(end ? { end } : {}),
      }
    : null
}

function readPPTXLineConnectionRef(
  connection: Element | null,
): PPTXLineConnectionRef | null {
  if (!connection) {
    return null
  }

  const objectId = connection.getAttribute('id')?.trim()

  return objectId
    ? {
        anchor: readPPTXLineConnectionAnchor(connection),
        objectId,
      }
    : null
}

function readPPTXLineConnectionAnchor(
  connection: Element,
): PPTLineConnection['anchor'] {
  const index = toPPTXNumber(connection.getAttribute('idx'))

  if (index === 0) {
    return 'left'
  }

  if (index === 1) {
    return 'top'
  }

  if (index === 2) {
    return 'right'
  }

  return index === 3 ? 'bottom' : 'center'
}

function readPPTXLineGeometry(spPr: Element | null): {
  end: PPTLine['end']
  geometry: PPTGeometry
  start: PPTLine['start']
} | null {
  const xfrm = spPr ? getDirectPPTXChildByLocalName(spPr, 'xfrm') : null
  const off = xfrm ? getDirectPPTXChildByLocalName(xfrm, 'off') : null
  const ext = xfrm ? getDirectPPTXChildByLocalName(xfrm, 'ext') : null
  const rawWidth = toPPTXNumber(ext?.getAttribute('cx'))
  const rawHeight = toPPTXNumber(ext?.getAttribute('cy'))

  if (!xfrm || rawWidth === null || rawHeight === null) {
    return null
  }

  let start = {
    x: emuToPx(toPPTXNumber(off?.getAttribute('x')) ?? 0),
    y: emuToPx(toPPTXNumber(off?.getAttribute('y')) ?? 0),
  }
  let end = {
    x: start.x + emuToPx(rawWidth),
    y: start.y + emuToPx(rawHeight),
  }

  if (isPPTXTrue(xfrm.getAttribute('flipH'))) {
    const startX = start.x
    start = { ...start, x: end.x }
    end = { ...end, x: startX }
  }

  if (isPPTXTrue(xfrm.getAttribute('flipV'))) {
    const startY = start.y
    start = { ...start, y: end.y }
    end = { ...end, y: startY }
  }

  const minX = Math.min(start.x, end.x)
  const minY = Math.min(start.y, end.y)
  const rawBounds = {
    h: Math.abs(end.y - start.y),
    w: Math.abs(end.x - start.x),
  }
  const geometry = {
    h: Math.max(24, rawBounds.h),
    ...(readPPTXRotation(xfrm) ?? {}),
    w: Math.max(24, rawBounds.w),
    x: minX - Math.max(0, 24 - rawBounds.w) / 2,
    y: minY - Math.max(0, 24 - rawBounds.h) / 2,
  }

  return {
    end: {
      x: end.x - geometry.x,
      y: end.y - geometry.y,
    },
    geometry,
    start: {
      x: start.x - geometry.x,
      y: start.y - geometry.y,
    },
  }
}

function readPPTXLineMarker(
  line: Element | null,
  marker: 'headEnd' | 'tailEnd',
): PPTLine['endMarker'] {
  const type = getDirectPPTXChildByLocalName(line, marker)
    ?.getAttribute('type')

  return type && type !== 'none' ? 'arrow' : 'none'
}

function readPPTXLineOpacity(line: Element | null) {
  const solidFill = getDirectPPTXChildByLocalName(line, 'solidFill')
  const opacity = solidFill ? readPPTXAlphaOpacity(solidFill) : null

  return opacity === null || opacity === 1 ? null : opacity
}

function readPPTXShapeElement(
  sp: Element,
  slideIndex: number,
  objectIndex: number,
  relationships: PPTXRelationshipMap,
  themeColors: PPTXThemeColorMap,
  themeFonts: PPTXThemeFontMap,
): PPTElement | null {
  const spPr = getDirectPPTXChildByLocalName(sp, 'spPr')
  const txBody = getDirectPPTXChildByLocalName(sp, 'txBody')
  const geometry = readPPTXElementGeometry(spPr)
  const textBody = readPPTXTextBody(txBody, themeColors)
  const stroke = readPPTXStroke(spPr, themeColors)
  const fill = readPPTXShapeFill(spPr, stroke, themeColors)
  const shadow = readPPTXElementShadow(spPr, themeColors)
  const hasPaint = fill !== null || stroke !== undefined
  const textAutoFit = readPPTXTextAutoFit(txBody)

  if (!geometry || (!textBody && !hasPaint)) {
    return null
  }

  const id = createPPTXImportedElementId(slideIndex, objectIndex)
  const name = readPPTXObjectName(sp, `Object ${objectIndex}`)
  const isTextBox = isPPTXTextBoxShape(sp) || (textBody !== null && !hasPaint)

  if (isTextBox) {
    return {
      ...(readPPTXElementAccessibility(sp) ?? {}),
      ...readPPTXElementFlip(spPr),
      geometry,
      ...(readPPTXElementHyperlink(sp, relationships) ?? {}),
      id,
      kind: 'textBox',
      ...(readPPTXElementLocked(sp) ? { locked: true } : {}),
      ...(readPPTXElementVisibility(sp) ?? {}),
      name,
      ...(shadow ? { shadow } : {}),
      style: readPPTXTextStyle(textBody, txBody, themeFonts),
      ...(textAutoFit ? { textAutoFit } : {}),
      textBody: textBody ?? { paragraphs: [] },
    }
  }

  return {
    ...(readPPTXElementAccessibility(sp) ?? {}),
    ...(readPPTXShapeCornerRadius(spPr, geometry) ?? {}),
    ...readPPTXElementFlip(spPr),
    ...(readPPTXElementHyperlink(sp, relationships) ?? {}),
    ...(stroke ? { stroke } : {}),
    ...(textBody ? {
      style: readPPTXTextStyle(textBody, txBody, themeFonts),
      ...(textAutoFit ? { textAutoFit } : {}),
      textBody,
    } : {}),
    fill: fill ?? { color: PPTX_DEFAULT_FILL_COLOR },
    geometry,
    id,
    kind: 'shape',
    ...(readPPTXElementLocked(sp) ? { locked: true } : {}),
    ...(readPPTXElementVisibility(sp) ?? {}),
    name,
    ...(shadow ? { shadow } : {}),
    shape: readPPTXShapeKind(spPr),
  }
}

async function readPPTXPictureElement({
  index,
  objectIndex,
  pic,
  relationships,
  slidePath,
  themeColors,
  zip,
}: {
  index: number
  objectIndex: number
  pic: Element
  relationships: PPTXRelationshipMap
  slidePath: string
  themeColors: PPTXThemeColorMap
  zip: JSZip
}): Promise<PPTImage | null> {
  const spPr = getDirectPPTXChildByLocalName(pic, 'spPr')
  const geometry = readPPTXElementGeometry(spPr)
  const blip = getFirstPPTXDescendantByLocalName(pic, 'blip')
  const mediaPath = readPPTXPictureMediaPath({
    blip,
    relationships,
    slidePath,
    zip,
  })
  const media = mediaPath ? zip.file(mediaPath) : null

  if (!geometry || !mediaPath || !media) {
    return null
  }

  const base64 = await media.async('base64')
  const mimeType = getPPTXMediaMimeType(mediaPath)
  const name = readPPTXObjectName(pic, `Image ${objectIndex}`)
  const altText = readPPTXObjectDescription(pic)
  const accessibility = readPPTXElementAccessibility(pic)
  const crop = readPPTXImageCrop(pic)
  const opacity = readPPTXImageOpacity(blip)
  const shadow = readPPTXElementShadow(spPr, themeColors)

  return {
    ...(accessibility ?? {}),
    alt: altText || name,
    ...(crop ? { crop } : {}),
    fit: crop ? 'cover' : 'contain',
    ...readPPTXElementFlip(spPr),
    geometry,
    ...(readPPTXElementHyperlink(pic, relationships) ?? {}),
    id: createPPTXImportedElementId(index, objectIndex),
    kind: 'image',
    ...(readPPTXElementLocked(pic) ? { locked: true } : {}),
    ...(readPPTXElementVisibility(pic) ?? {}),
    name,
    ...(opacity === null ? {} : { opacity }),
    ...(shadow ? { shadow } : {}),
    src: `data:${mimeType};base64,${base64}`,
  }
}

function readPPTXPictureMediaPath({
  blip,
  relationships,
  slidePath,
  zip,
}: {
  blip: Element | null
  relationships: PPTXRelationshipMap
  slidePath: string
  zip: JSZip
}) {
  const svgBlip = getFirstPPTXDescendantByLocalName(blip, 'svgBlip')
  const relationshipIds = [
    readPPTXEmbedRelationshipId(svgBlip),
    readPPTXEmbedRelationshipId(blip),
  ].filter((id): id is string => id !== null)
  const seen = new Set<string>()

  for (const relationshipId of relationshipIds) {
    if (seen.has(relationshipId)) {
      continue
    }

    seen.add(relationshipId)

    const relationship = relationships.get(relationshipId)
    const mediaPath = relationship
      ? resolvePPTXRelationshipTarget(slidePath, relationship.target)
      : null

    if (mediaPath && zip.file(mediaPath)) {
      return mediaPath
    }
  }

  return null
}

function readPPTXEmbedRelationshipId(element: Element | null) {
  return element?.getAttribute('r:embed') ??
    element?.getAttribute('embed') ??
    null
}

function readPPTXImageCrop(pic: Element): PPTImage['crop'] | null {
  const srcRect = getFirstPPTXDescendantByLocalName(pic, 'srcRect')

  if (!srcRect) {
    return null
  }

  const left = toPPTXNumber(srcRect.getAttribute('l')) ?? 0
  const right = toPPTXNumber(srcRect.getAttribute('r')) ?? 0
  const top = toPPTXNumber(srcRect.getAttribute('t')) ?? 0
  const bottom = toPPTXNumber(srcRect.getAttribute('b')) ?? 0
  const x = clampPPTXPercent(50 + (left - right) / 2_000)
  const y = clampPPTXPercent(50 + (top - bottom) / 2_000)

  return x === 50 && y === 50 ? null : { x, y }
}

function readPPTXTableElement(
  graphicFrame: Element,
  slideIndex: number,
  objectIndex: number,
  relationships: PPTXRelationshipMap,
  themeColors: PPTXThemeColorMap,
): PPTElement | null {
  const table = getFirstPPTXDescendantByLocalName(graphicFrame, 'tbl')
  const geometry = readPPTXElementGeometry(graphicFrame)
  const shadow = readPPTXElementShadow(graphicFrame, themeColors)
  const cellRows = table
    ? getDirectPPTXChildrenByLocalName(table, 'tr')
      .map((row) => getDirectPPTXChildrenByLocalName(row, 'tc'))
      .filter((row) => row.length > 0)
    : []
  const rows = cellRows.map((row) =>
    row.map((cell) => readPPTXPlainTextBody(cell).trim()))

  if (!geometry || rows.length === 0) {
    return null
  }

  const cellStyles = readPPTXTableCellStyles(cellRows, themeColors)
  const columnWidths = table
    ? readPPTXTableColumnWidths(table, getPPTTableColumnCount(rows))
    : undefined
  const rowHeights = table
    ? readPPTXTableRowHeights(table, rows.length)
    : undefined

  return {
    ...(readPPTXElementAccessibility(graphicFrame) ?? {}),
    ...(cellStyles ? { cellStyles } : {}),
    ...(columnWidths ? { columnWidths } : {}),
    ...readPPTXElementFlip(graphicFrame),
    geometry,
    ...(readPPTXElementHyperlink(graphicFrame, relationships) ?? {}),
    id: createPPTXImportedElementId(slideIndex, objectIndex),
    kind: 'table',
    ...(readPPTXElementLocked(graphicFrame) ? { locked: true } : {}),
    ...(readPPTXElementVisibility(graphicFrame) ?? {}),
    name: readPPTXObjectName(graphicFrame, `Table ${objectIndex}`),
    ...(rowHeights ? { rowHeights } : {}),
    ...(shadow ? { shadow } : {}),
    rows,
  }
}

function readPPTXTableCellStyles(
  cellRows: readonly (readonly Element[])[],
  themeColors: PPTXThemeColorMap,
): PPTTable['cellStyles'] {
  const styles = cellRows.map((row) =>
    row.map((cell): PPTTableCellStyle => {
      const borders = readPPTXTableCellBorders(cell, themeColors)
      const fill = readPPTXTableCellFill(cell, themeColors)
      const span = readPPTXTableCellSpan(cell)
      const textStyle = readPPTXTableCellTextStyle(cell, themeColors)

      return {
        ...(borders ? { borders } : {}),
        ...(fill ? { fill } : {}),
        ...span,
        ...(textStyle ? { textStyle } : {}),
      }
    }))

  return styles.some((row) =>
    row.some((style) =>
      style.borders ||
      style.colSpan ||
      style.fill ||
      style.hidden ||
      style.rowSpan ||
      style.textStyle))
    ? styles
    : undefined
}

function readPPTXTableCellSpan(cell: Element): Pick<
  PPTTableCellStyle,
  'colSpan' | 'hidden' | 'rowSpan'
> {
  const colSpan = readPPTXTableCellSpanValue(cell, 'gridSpan')
  const rowSpan = readPPTXTableCellSpanValue(cell, 'rowSpan')
  const hidden = isPPTXTrue(cell.getAttribute('hMerge')) ||
    isPPTXTrue(cell.getAttribute('vMerge'))

  return {
    ...(colSpan > 1 ? { colSpan } : {}),
    ...(hidden ? { hidden: true } : {}),
    ...(rowSpan > 1 ? { rowSpan } : {}),
  }
}

function readPPTXTableCellSpanValue(
  cell: Element,
  attribute: 'gridSpan' | 'rowSpan',
) {
  const value = toPPTXPositiveNumber(cell.getAttribute(attribute))

  return value === null ? 1 : Math.max(1, Math.floor(value))
}

function readPPTXTableCellBorders(
  cell: Element,
  themeColors: PPTXThemeColorMap,
): PPTTableCellBorders | undefined {
  const tcPr = getDirectPPTXChildByLocalName(cell, 'tcPr')
  const borders = {
    bottom: readPPTXTableCellBorderSide(tcPr, 'lnB', themeColors),
    left: readPPTXTableCellBorderSide(tcPr, 'lnL', themeColors),
    right: readPPTXTableCellBorderSide(tcPr, 'lnR', themeColors),
    top: readPPTXTableCellBorderSide(tcPr, 'lnT', themeColors),
  }

  return borders.bottom || borders.left || borders.right || borders.top
    ? borders
    : undefined
}

function readPPTXTableCellBorderSide(
  tcPr: Element | null,
  tagName: 'lnB' | 'lnL' | 'lnR' | 'lnT',
  themeColors: PPTXThemeColorMap,
) {
  return readPPTXStrokeLine(
    getDirectPPTXChildByLocalName(tcPr, tagName),
    themeColors,
  )
}

function readPPTXTableCellFill(
  cell: Element,
  themeColors: PPTXThemeColorMap,
): PPTFill | undefined {
  const tcPr = getDirectPPTXChildByLocalName(cell, 'tcPr')

  return readPPTXFill(tcPr, themeColors) ?? undefined
}

function readPPTXTableCellTextStyle(
  cell: Element,
  themeColors: PPTXThemeColorMap,
): PPTTableCellTextStyle | undefined {
  const tcPr = getDirectPPTXChildByLocalName(cell, 'tcPr')
  const txBody = getDirectPPTXChildByLocalName(cell, 'txBody')
  const textBody = readPPTXTextBody(txBody, themeColors)
  const textInset = readPPTXTextInset(tcPr)
  const verticalAlign = readPPTXTableCellVerticalAlign(tcPr)
  const firstParagraph = textBody?.paragraphs
    .find((paragraph) =>
      paragraph.runs.some((run) => run.text.trim().length > 0)) ??
    textBody?.paragraphs[0]
  const firstRun = firstParagraph?.runs
    .find((run) => run.text.trim().length > 0) ??
    firstParagraph?.runs[0]

  const textStyle = {
    ...(firstParagraph?.align ? { align: firstParagraph.align } : {}),
    ...(firstRun?.color ? { color: firstRun.color } : {}),
    ...(firstRun?.size === undefined ? {} : { fontSize: firstRun.size }),
    ...(firstRun?.bold === true ? { fontWeight: 'bold' as const } : {}),
    ...(textInset ? { textInset } : {}),
    ...(verticalAlign ? { verticalAlign } : {}),
  }

  return Object.keys(textStyle).length > 0 ? textStyle : undefined
}

function readPPTXTableCellVerticalAlign(
  tcPr: Element | null,
): PPTTableCellTextStyle['verticalAlign'] | undefined {
  const anchor = tcPr?.getAttribute('anchor')

  if (anchor === 'ctr') {
    return 'middle'
  }

  if (anchor === 'b') {
    return 'bottom'
  }

  return anchor === 't' ? 'top' : undefined
}

function readPPTXTableColumnWidths(
  table: Element,
  columnCount: number,
): PPTTable['columnWidths'] {
  const grid = getDirectPPTXChildByLocalName(table, 'tblGrid')
  const widths = grid
    ? getDirectPPTXChildrenByLocalName(grid, 'gridCol')
      .map((column) => toPPTXPositiveNumber(column.getAttribute('w')))
    : []
  const visibleWidths = widths.slice(0, columnCount)

  return visibleWidths.length === columnCount &&
    visibleWidths.every((width) => width !== null && width > 0)
    ? visibleWidths.map((width) => emuToPx(width ?? 0))
    : undefined
}

function readPPTXTableRowHeights(
  table: Element,
  rowCount: number,
): PPTTable['rowHeights'] {
  const heights = getDirectPPTXChildrenByLocalName(table, 'tr')
    .map((row) => toPPTXPositiveNumber(row.getAttribute('h')))
  const visibleHeights = heights.slice(0, rowCount)

  return visibleHeights.length === rowCount &&
    visibleHeights.every((height) => height !== null && height > 0)
    ? visibleHeights.map((height) => emuToPx(height ?? 0))
    : undefined
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

function readPPTXGroupTransform(group: Element): PPTXGroupTransform {
  const grpSpPr = getDirectPPTXChildByLocalName(group, 'grpSpPr')
  const xfrm = getDirectPPTXChildByLocalName(grpSpPr, 'xfrm')
  const off = getDirectPPTXChildByLocalName(xfrm, 'off')
  const ext = getDirectPPTXChildByLocalName(xfrm, 'ext')
  const childOff = getDirectPPTXChildByLocalName(xfrm, 'chOff')
  const childExt = getDirectPPTXChildByLocalName(xfrm, 'chExt')
  const rawWidth = toPPTXPositiveNumber(ext?.getAttribute('cx'))
  const rawHeight = toPPTXPositiveNumber(ext?.getAttribute('cy'))
  const rawChildWidth = toPPTXPositiveNumber(childExt?.getAttribute('cx'))
  const rawChildHeight = toPPTXPositiveNumber(childExt?.getAttribute('cy'))

  return {
    childOffsetX: emuToPx(toPPTXNumber(childOff?.getAttribute('x')) ?? 0),
    childOffsetY: emuToPx(toPPTXNumber(childOff?.getAttribute('y')) ?? 0),
    offsetX: emuToPx(toPPTXNumber(off?.getAttribute('x')) ?? 0),
    offsetY: emuToPx(toPPTXNumber(off?.getAttribute('y')) ?? 0),
    scaleX: rawWidth !== null && rawChildWidth !== null && rawChildWidth > 0
      ? rawWidth / rawChildWidth
      : 1,
    scaleY: rawHeight !== null && rawChildHeight !== null && rawChildHeight > 0
      ? rawHeight / rawChildHeight
      : 1,
  }
}

function composePPTXGroupTransforms(
  parent: PPTXGroupTransform,
  child: PPTXGroupTransform,
): PPTXGroupTransform {
  return {
    childOffsetX: child.childOffsetX,
    childOffsetY: child.childOffsetY,
    offsetX: transformPPTXGroupCoordinate(
      child.offsetX,
      parent.offsetX,
      parent.childOffsetX,
      parent.scaleX,
    ),
    offsetY: transformPPTXGroupCoordinate(
      child.offsetY,
      parent.offsetY,
      parent.childOffsetY,
      parent.scaleY,
    ),
    scaleX: parent.scaleX * child.scaleX,
    scaleY: parent.scaleY * child.scaleY,
  }
}

function transformPPTXElement(
  element: PPTElement,
  transform: PPTXGroupTransform,
): PPTElement {
  const geometry = transformPPTXGeometry(element.geometry, transform)

  if (element.kind === 'line') {
    return {
      ...element,
      end: transformPPTXLinePoint(element.end, transform),
      geometry,
      start: transformPPTXLinePoint(element.start, transform),
    }
  }

  if (element.kind === 'table') {
    return {
      ...element,
      ...(element.columnWidths
        ? { columnWidths: transformPPTXTableTrackSizes(element.columnWidths, transform.scaleX) }
        : {}),
      geometry,
      ...(element.rowHeights
        ? { rowHeights: transformPPTXTableTrackSizes(element.rowHeights, transform.scaleY) }
        : {}),
    }
  }

  return {
    ...element,
    geometry,
  }
}

function transformPPTXTableTrackSizes(
  trackSizes: readonly number[],
  scale: number,
) {
  return trackSizes.map((size) => Math.max(1, Math.round(size * scale)))
}

function transformPPTXGeometry(
  geometry: PPTGeometry,
  transform: PPTXGroupTransform,
): PPTGeometry {
  return {
    ...geometry,
    h: Math.max(1, Math.round(geometry.h * transform.scaleY)),
    w: Math.max(1, Math.round(geometry.w * transform.scaleX)),
    x: transformPPTXGroupCoordinate(
      geometry.x,
      transform.offsetX,
      transform.childOffsetX,
      transform.scaleX,
    ),
    y: transformPPTXGroupCoordinate(
      geometry.y,
      transform.offsetY,
      transform.childOffsetY,
      transform.scaleY,
    ),
  }
}

function transformPPTXLinePoint(
  point: PPTLine['start'],
  transform: PPTXGroupTransform,
) {
  return {
    x: Math.round(point.x * transform.scaleX),
    y: Math.round(point.y * transform.scaleY),
  }
}

function transformPPTXGroupCoordinate(
  value: number,
  offset: number,
  childOffset: number,
  scale: number,
) {
  return Math.round(offset + (value - childOffset) * scale)
}

function isPPTXIdentityGroupTransform(transform: PPTXGroupTransform) {
  return transform.offsetX === 0 &&
    transform.offsetY === 0 &&
    transform.childOffsetX === 0 &&
    transform.childOffsetY === 0 &&
    transform.scaleX === 1 &&
    transform.scaleY === 1
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

  return {
    h: emuToPx(height),
    ...(readPPTXRotation(xfrm) ?? {}),
    w: emuToPx(width),
    x: emuToPx(toPPTXNumber(off?.getAttribute('x')) ?? 0),
    y: emuToPx(toPPTXNumber(off?.getAttribute('y')) ?? 0),
  }
}

function readPPTXElementFlip(container: Element | null) {
  const xfrm = container ? getDirectPPTXChildByLocalName(container, 'xfrm') : null

  return {
    ...(isPPTXTrue(xfrm?.getAttribute('flipH')) ? { flipH: true } : {}),
    ...(isPPTXTrue(xfrm?.getAttribute('flipV')) ? { flipV: true } : {}),
  }
}

function readPPTXElementShadow(
  container: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTElementShadow | null {
  const outerShadow = getFirstPPTXDescendantByLocalName(container, 'outerShdw')

  if (!outerShadow) {
    return null
  }

  const direction = toPPTXNumber(outerShadow.getAttribute('dir'))

  return {
    angle: direction === null ? 45 : normalizePPTXAngle(direction / 60_000),
    blur: emuToPx(toPPTXPositiveNumber(outerShadow.getAttribute('blurRad')) ?? 0),
    color: readPPTXColor(outerShadow, themeColors) ?? '#000000',
    distance: emuToPx(toPPTXPositiveNumber(outerShadow.getAttribute('dist')) ?? 0),
    opacity: readPPTXAlphaOpacity(outerShadow) ?? 1,
  }
}

function readPPTXElementAccessibility(element: Element) {
  const altText = readPPTXObjectDescription(element)

  return altText ? { accessibility: { altText } } : null
}

function readPPTXElementLocked(element: Element) {
  return PPTX_LOCK_TAG_NAMES.some((tagName) =>
    getPPTXDescendantsByLocalName(element, tagName)
      .some((locks) =>
        PPTX_LOCK_ATTRIBUTE_NAMES.some((attribute) =>
          isPPTXTrue(locks.getAttribute(attribute)))))
}

function readPPTXElementVisibility(element: Element) {
  const hidden = getFirstPPTXDescendantByLocalName(element, 'cNvPr')
    ?.getAttribute('hidden')

  return isPPTXTrue(hidden) ? { visible: false } : null
}

function readPPTXRotation(xfrm: Element) {
  const rotation = toPPTXNumber(xfrm.getAttribute('rot'))

  return rotation === null ? null : { rotation: rotation / 60_000 }
}

function readPPTXTextBody(
  txBody: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTTextBody | null {
  if (!txBody) {
    return null
  }

  const listStyle = getDirectPPTXChildByLocalName(txBody, 'lstStyle')
  const paragraphs = getDirectPPTXChildrenByLocalName(txBody, 'p')
    .map((paragraph) => readPPTXParagraph(paragraph, listStyle, themeColors))
  const hasText = paragraphs.some((paragraph) =>
    paragraph.runs.some((run) => run.text.length > 0))

  return hasText || paragraphs.length > 0 ? { paragraphs } : null
}

function readPPTXPlainTextBody(root: Document | Element) {
  return getPPTXDescendantsByLocalName(root, 'p')
    .map(readPPTXPlainParagraphText)
    .filter((text) => text.length > 0)
    .join('\n')
}

function readPPTXPlainParagraphText(paragraph: Element) {
  return Array.from(paragraph.children)
    .map((child) => {
      if (child.localName === 'br') {
        return '\n'
      }

      if (child.localName === 'r' || child.localName === 'fld') {
        return getFirstPPTXDescendantByLocalName(child, 't')?.textContent ?? ''
      }

      return ''
    })
    .join('')
}

function readPPTXParagraph(
  paragraph: Element,
  listStyle: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTParagraph {
  const pPr = getDirectPPTXChildByLocalName(paragraph, 'pPr')
  const level = readPPTXParagraphLevel(pPr)
  const listStylePPr = readPPTXTextListStyleParagraphProperties(
    listStyle,
    level ?? 0,
  )
  const defaultRunProperties = readPPTXParagraphDefaultRunProperties(
    paragraph,
    pPr,
    listStylePPr,
  )
  const align = readPPTXParagraphAlign(pPr) ??
    readPPTXParagraphAlign(listStylePPr)
  const bullet = readPPTXParagraphBullet(pPr, listStylePPr)
  const spacing = readPPTXParagraphSpacing(
    pPr,
    listStylePPr,
    defaultRunProperties,
  )
  const runs = Array.from(paragraph.children)
    .flatMap((child) =>
      readPPTXTextRun(child, defaultRunProperties, themeColors))

  return {
    ...(align ? { align } : {}),
    ...(bullet ? { bullet } : {}),
    ...(level === undefined ? {} : { level }),
    ...spacing,
    runs: runs.length > 0 ? runs : [{ text: '' }],
  }
}

function readPPTXParagraphDefaultRunProperties(
  paragraph: Element,
  pPr: Element | null,
  listStylePPr: Element | null,
) {
  return getDirectPPTXChildByLocalName(pPr, 'defRPr') ??
    getDirectPPTXChildByLocalName(listStylePPr, 'defRPr') ??
    getDirectPPTXChildByLocalName(paragraph, 'endParaRPr')
}

function readPPTXTextListStyleParagraphProperties(
  listStyle: Element | null,
  level: number,
) {
  if (!listStyle) {
    return null
  }

  const clampedLevel = Math.max(0, Math.min(8, Math.floor(level)))

  return getDirectPPTXChildByLocalName(listStyle, `lvl${clampedLevel + 1}pPr`) ??
    getDirectPPTXChildByLocalName(listStyle, 'defPPr')
}

function readPPTXTextRun(
  node: Element,
  defaultRunProperties: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTRun[] {
  if (node.localName !== 'r' && node.localName !== 'fld' && node.localName !== 'br') {
    return []
  }

  const rPr = getDirectPPTXChildByLocalName(node, 'rPr')
  const style = readPPTXTextRunStyle(rPr, defaultRunProperties, themeColors)

  if (node.localName === 'br') {
    return [{ ...style, text: '\n' }]
  }

  const text = getFirstPPTXDescendantByLocalName(node, 't')?.textContent ?? ''

  return [{ ...style, text }]
}

function readPPTXTextRunStyle(
  rPr: Element | null,
  defaultRunProperties: Element | null,
  themeColors: PPTXThemeColorMap,
): Omit<PPTRun, 'text'> {
  const color = readPPTXRunColor(rPr, defaultRunProperties, themeColors)
  const highlight = readPPTXRunHighlight(rPr, themeColors) ??
    readPPTXRunHighlight(defaultRunProperties, themeColors)
  const size = readPPTXRunSize(rPr, defaultRunProperties)

  return {
    ...(readPPTXRunBooleanAttribute(rPr, defaultRunProperties, 'b')
      ? { bold: true }
      : {}),
    ...(color ? { color } : {}),
    ...(highlight ? { highlight } : {}),
    ...(readPPTXRunBooleanAttribute(rPr, defaultRunProperties, 'i')
      ? { italic: true }
      : {}),
    ...(size === null ? {} : { size: textSizeToPx(size) }),
    ...(readPPTXRunStrikethrough(rPr, defaultRunProperties)
      ? { strikethrough: true }
      : {}),
    ...(readPPTXRunUnderline(rPr, defaultRunProperties)
      ? { underline: true }
      : {}),
  }
}

function readPPTXRunColor(
  rPr: Element | null,
  defaultRunProperties: Element | null,
  themeColors: PPTXThemeColorMap,
) {
  return readPPTXFill(rPr, themeColors)?.color ??
    readPPTXFill(defaultRunProperties, themeColors)?.color
}

function readPPTXRunSize(
  rPr: Element | null,
  defaultRunProperties: Element | null,
) {
  return toPPTXPositiveNumber(rPr?.getAttribute('sz')) ??
    toPPTXPositiveNumber(defaultRunProperties?.getAttribute('sz'))
}

function readPPTXRunBooleanAttribute(
  rPr: Element | null,
  defaultRunProperties: Element | null,
  attribute: 'b' | 'i',
) {
  const own = rPr?.getAttribute(attribute)

  return own === null || own === undefined
    ? isPPTXTrue(defaultRunProperties?.getAttribute(attribute))
    : isPPTXTrue(own)
}

function readPPTXRunUnderline(
  rPr: Element | null,
  defaultRunProperties: Element | null,
) {
  return hasPPTXAttribute(rPr, 'u')
    ? readPPTXUnderline(rPr)
    : readPPTXUnderline(defaultRunProperties)
}

function readPPTXRunStrikethrough(
  rPr: Element | null,
  defaultRunProperties: Element | null,
) {
  return hasPPTXAttribute(rPr, 'strike')
    ? readPPTXStrikethrough(rPr)
    : readPPTXStrikethrough(defaultRunProperties)
}

function readPPTXTextStyle(
  textBody: PPTTextBody | null,
  txBody: Element | null,
  themeFonts: PPTXThemeFontMap,
): PPTTextStyle {
  const firstRun = textBody?.paragraphs
    .flatMap((paragraph) => paragraph.runs)
    .find((run) => run.text.trim().length > 0) ??
    textBody?.paragraphs[0]?.runs[0]
  const fontFamily = readPPTXFirstTypeface(txBody, themeFonts)

  return {
    color: firstRun?.color ?? PPTX_DEFAULT_TEXT_COLOR,
    ...(fontFamily ? { fontFamily } : {}),
    fontSize: firstRun?.size ?? PPTX_DEFAULT_TEXT_SIZE,
    ...(firstRun?.bold === true ? { fontWeight: 'bold' } : {}),
    ...readPPTXTextFrameStyle(txBody),
  }
}

function readPPTXTextFrameStyle(
  txBody: Element | null,
): Pick<PPTTextStyle, 'textInset' | 'verticalAlign'> {
  const bodyPr = getDirectPPTXChildByLocalName(txBody, 'bodyPr')
  const verticalAlign = readPPTXTextVerticalAlign(bodyPr)
  const textInset = readPPTXTextInset(bodyPr)

  return {
    ...(textInset ? { textInset } : {}),
    ...(verticalAlign ? { verticalAlign } : {}),
  }
}

function readPPTXTextAutoFit(
  txBody: Element | null,
): PPTTextAutoFit | undefined {
  const bodyPr = getDirectPPTXChildByLocalName(txBody, 'bodyPr')

  return getDirectPPTXChildByLocalName(bodyPr, 'spAutoFit')
    ? 'resizeShapeToFitText'
    : undefined
}

function readPPTXTextVerticalAlign(
  bodyPr: Element | null,
): PPTTextStyle['verticalAlign'] | undefined {
  const anchor = bodyPr?.getAttribute('anchor')

  if (anchor === 'ctr') {
    return 'middle'
  }

  if (anchor === 'b') {
    return 'bottom'
  }

  return anchor === 't' ? 'top' : undefined
}

function readPPTXTextInset(
  bodyPr: Element | null,
): PPTTextStyle['textInset'] | undefined {
  if (!bodyPr) {
    return undefined
  }

  const top = readPPTXTextInsetSide(bodyPr, 'tIns', 'marT')
  const right = readPPTXTextInsetSide(bodyPr, 'rIns', 'marR')
  const bottom = readPPTXTextInsetSide(bodyPr, 'bIns', 'marB')
  const left = readPPTXTextInsetSide(bodyPr, 'lIns', 'marL')

  return top === undefined &&
    right === undefined &&
    bottom === undefined &&
    left === undefined
    ? undefined
    : {
        bottom: bottom ?? 0,
        left: left ?? 0,
        right: right ?? 0,
        top: top ?? 0,
      }
}

function readPPTXTextInsetSide(
  bodyPr: Element,
  insetAttribute: 'bIns' | 'lIns' | 'rIns' | 'tIns',
  legacyMarginAttribute: 'marB' | 'marL' | 'marR' | 'marT',
) {
  const value = toPPTXPositiveNumber(
    bodyPr.getAttribute(insetAttribute) ??
      bodyPr.getAttribute(legacyMarginAttribute),
  )

  return value === null ? undefined : emuToPx(value)
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
  fallbackPPr: Element | null,
): PPTParagraph['bullet'] | undefined {
  if (getDirectPPTXChildByLocalName(pPr, 'buNone')) {
    return undefined
  }

  const bullet = readPPTXParagraphBulletFromProperties(pPr)

  if (bullet) {
    return bullet
  }

  if (getDirectPPTXChildByLocalName(fallbackPPr, 'buNone')) {
    return undefined
  }

  return readPPTXParagraphBulletFromProperties(fallbackPPr)
}

function readPPTXParagraphBulletFromProperties(
  pPr: Element | null,
): PPTParagraph['bullet'] | undefined {
  if (!pPr) {
    return undefined
  }

  if (getDirectPPTXChildByLocalName(pPr, 'buAutoNum')) {
    return 'numbered'
  }

  return getDirectPPTXChildByLocalName(pPr, 'buChar') ||
    getDirectPPTXChildByLocalName(pPr, 'buBlip')
    ? 'bullet'
    : undefined
}

function readPPTXParagraphLevel(pPr: Element | null) {
  const level = toPPTXNumber(pPr?.getAttribute('lvl'))

  return level === null ? undefined : Math.max(0, level)
}

function readPPTXParagraphSpacing(
  pPr: Element | null,
  fallbackPPr: Element | null,
  defaultRunProperties: Element | null,
): Pick<PPTParagraph, 'lineHeight' | 'spacingAfter' | 'spacingBefore'> {
  const lineHeight = readPPTXParagraphLineHeight(pPr, defaultRunProperties) ??
    readPPTXParagraphLineHeight(fallbackPPr, defaultRunProperties)
  const spacingBefore = readPPTXParagraphSpacingPixels(pPr, 'spcBef') ??
    readPPTXParagraphSpacingPixels(fallbackPPr, 'spcBef')
  const spacingAfter = readPPTXParagraphSpacingPixels(pPr, 'spcAft') ??
    readPPTXParagraphSpacingPixels(fallbackPPr, 'spcAft')

  return {
    ...(lineHeight === undefined ? {} : { lineHeight }),
    ...(spacingAfter === undefined ? {} : { spacingAfter }),
    ...(spacingBefore === undefined ? {} : { spacingBefore }),
  }
}

function readPPTXParagraphLineHeight(
  pPr: Element | null,
  defaultRunProperties: Element | null,
) {
  const spacing = getDirectPPTXChildByLocalName(pPr, 'lnSpc')
  const percent = getDirectPPTXChildByLocalName(spacing, 'spcPct')
  const percentValue = toPPTXPositiveNumber(percent?.getAttribute('val'))

  if (percentValue !== null) {
    return percentValue / 100_000
  }

  const points = getDirectPPTXChildByLocalName(spacing, 'spcPts')
  const pointValue = toPPTXPositiveNumber(points?.getAttribute('val'))

  if (pointValue === null) {
    return undefined
  }

  const exactPoints = pointValue / 100
  const fontSizeUnits = readPPTXRunSize(defaultRunProperties, null)
  const fontPoints = fontSizeUnits === null
    ? PPTX_DEFAULT_TEXT_SIZE * PPTX_POINTS_PER_PIXEL
    : fontSizeUnits / PPTX_TEXT_SIZE_UNITS_PER_POINT

  return fontPoints > 0
    ? Math.round((exactPoints / fontPoints) * 1000) / 1000
    : undefined
}

function readPPTXParagraphSpacingPixels(
  pPr: Element | null,
  localName: 'spcAft' | 'spcBef',
) {
  const spacing = getDirectPPTXChildByLocalName(pPr, localName)
  const points = getDirectPPTXChildByLocalName(spacing, 'spcPts')
  const value = toPPTXPositiveNumber(points?.getAttribute('val'))

  return value === null ? undefined : pointToPx(value / 100)
}

function readPPTXUnderline(rPr: Element | null) {
  const underline = rPr?.getAttribute('u')

  return underline !== null && underline !== undefined && underline !== 'none'
}

function readPPTXStrikethrough(rPr: Element | null) {
  const strike = rPr?.getAttribute('strike')

  return strike !== null &&
    strike !== undefined &&
    strike !== 'noStrike' &&
    strike !== 'none'
}

function readPPTXRunHighlight(
  rPr: Element | null,
  themeColors: PPTXThemeColorMap,
) {
  const highlight = getDirectPPTXChildByLocalName(rPr, 'highlight')

  return highlight ? readPPTXColor(highlight, themeColors) : undefined
}

function readPPTXTypeface(
  rPr: Element | null,
  themeFonts: PPTXThemeFontMap,
) {
  for (const localName of ['latin', 'ea', 'cs']) {
    const typeface = getDirectPPTXChildByLocalName(rPr, localName)
      ?.getAttribute('typeface')
      ?.trim()

    if (!typeface) {
      continue
    }

    if (!typeface.startsWith('+')) {
      return typeface
    }

    const themeTypeface = themeFonts[typeface]

    if (themeTypeface) {
      return themeTypeface
    }
  }

  return undefined
}

function readPPTXFirstTypeface(
  txBody: Element | null,
  themeFonts: PPTXThemeFontMap,
) {
  if (!txBody) {
    return undefined
  }

  for (const properties of Array.from(txBody.getElementsByTagName('*'))) {
    if (properties.localName !== 'rPr' &&
      properties.localName !== 'defRPr' &&
      properties.localName !== 'endParaRPr') {
      continue
    }

    const typeface = readPPTXTypeface(properties, themeFonts)

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

function readPPTXShapeCornerRadius(
  spPr: Element | null,
  geometry: PPTGeometry,
) {
  const presetGeometry = getFirstPPTXDescendantByLocalName(spPr, 'prstGeom')
  const preset = presetGeometry
    ?.getAttribute('prst')

  if (preset !== 'roundRect') {
    return null
  }

  const adjust = readPPTXPresetGeometryAdjust(
    presetGeometry,
    'adj',
    PPTX_ROUND_RECT_DEFAULT_ADJUST,
  )
  const radius = Math.round(
    Math.min(geometry.w, geometry.h) * adjust / 100_000,
  )

  return { cornerRadius: Math.max(0, radius) }
}

function readPPTXPresetGeometryAdjust(
  presetGeometry: Element | null,
  name: string,
  fallback: number,
) {
  const formula = presetGeometry
    ? getPPTXDescendantsByLocalName(presetGeometry, 'gd')
      .find((guide) => guide.getAttribute('name') === name)
      ?.getAttribute('fmla') ?? ''
    : ''
  const value = toPPTXNumber(formula.match(/^val\s+(-?\d+(?:\.\d+)?)$/)?.[1])
  const normalized = value === null ? fallback : value

  return Math.min(
    PPTX_ROUND_RECT_MAX_ADJUST,
    Math.max(0, normalized),
  )
}

function readPPTXShapeFill(
  spPr: Element | null,
  stroke: PPTStroke | undefined,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  const fill = readPPTXFill(spPr, themeColors)

  if (fill || !hasPPTXNoFill(spPr) || !stroke) {
    return fill
  }

  return {
    color: PPTX_DEFAULT_FILL_COLOR,
    opacity: 0,
  }
}

function readPPTXFill(
  container: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  if (!container || hasPPTXNoFill(container)) {
    return null
  }

  return readPPTXSolidFill(container, themeColors) ??
    readPPTXGradientFill(container, themeColors) ??
    readPPTXPatternFill(container, themeColors)
}

function readPPTXSolidFill(
  container: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  const solidFill = getDirectPPTXChildByLocalName(container, 'solidFill')

  return solidFill ? readPPTXColorFill(solidFill, themeColors) : null
}

function readPPTXGradientFill(
  container: Element,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  const gradFill = getDirectPPTXChildByLocalName(container, 'gradFill')
  const stop = gradFill
    ? getPPTXDescendantsByLocalName(gradFill, 'gs')
      .sort(comparePPTXGradientStopPositions)
      .find((gradientStop) => readPPTXColor(gradientStop, themeColors))
    : null

  return stop ? readPPTXColorFill(stop, themeColors) : null
}

function readPPTXPatternFill(
  container: Element,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  const pattFill = getDirectPPTXChildByLocalName(container, 'pattFill')

  if (!pattFill) {
    return null
  }

  const foreground = getDirectPPTXChildByLocalName(pattFill, 'fgClr')
  const background = getDirectPPTXChildByLocalName(pattFill, 'bgClr')

  return readPPTXColorFill(foreground, themeColors) ??
    readPPTXColorFill(background, themeColors)
}

function readPPTXColorFill(
  colorContainer: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTFill | null {
  const color = colorContainer
    ? readPPTXColor(colorContainer, themeColors)
    : undefined

  if (!color || !colorContainer) {
    return null
  }

  const opacity = readPPTXAlphaOpacity(colorContainer)

  return opacity === null ? { color } : { color, opacity }
}

function readPPTXStroke(
  spPr: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTStroke | undefined {
  const line = spPr ? getDirectPPTXChildByLocalName(spPr, 'ln') : null

  return readPPTXStrokeLine(line, themeColors)
}

function readPPTXStrokeLine(
  line: Element | null,
  themeColors: PPTXThemeColorMap,
): PPTStroke | undefined {
  if (!line || hasPPTXNoFill(line)) {
    return undefined
  }

  const dash = readPPTXStrokeDash(line)

  return {
    ...(dash ? { dash } : {}),
    color: readPPTXFill(line, themeColors)?.color ?? PPTX_DEFAULT_STROKE_COLOR,
    width: Math.max(1, emuToPx(toPPTXPositiveNumber(line.getAttribute('w')) ?? PPTX_EMUS_PER_PIXEL)),
  }
}

function comparePPTXGradientStopPositions(left: Element, right: Element) {
  return (toPPTXPositiveNumber(left.getAttribute('pos')) ?? 0) -
    (toPPTXPositiveNumber(right.getAttribute('pos')) ?? 0)
}

function readPPTXStrokeDash(line: Element): PPTStroke['dash'] | undefined {
  const value = getDirectPPTXChildByLocalName(line, 'prstDash')
    ?.getAttribute('val')

  if (value === 'dot' || value === 'sysDot') {
    return 'dot'
  }

  return value === 'dash' || value === 'lgDash' || value === 'sysDash'
    ? 'dash'
    : undefined
}

function readPPTXColor(
  solidFill: Element,
  themeColors: PPTXThemeColorMap,
) {
  const srgbColor = getDirectPPTXChildByLocalName(solidFill, 'srgbClr')
  const schemeColor = getDirectPPTXChildByLocalName(solidFill, 'schemeClr')
  const presetColor = getDirectPPTXChildByLocalName(solidFill, 'prstClr')
  const systemColor = getDirectPPTXChildByLocalName(solidFill, 'sysClr')
  const color = [
    {
      color: readPPTXHexColor(srgbColor?.getAttribute('val')),
      element: srgbColor,
    },
    {
      color: readPPTXSchemeColor(schemeColor?.getAttribute('val'), themeColors),
      element: schemeColor,
    },
    {
      color: readPPTXPresetColor(presetColor?.getAttribute('val')),
      element: presetColor,
    },
    {
      color: readPPTXHexColor(systemColor?.getAttribute('lastClr')) ??
        readPPTXSystemColor(systemColor?.getAttribute('val')),
      element: systemColor,
    },
  ].find((candidate) => candidate.color && candidate.element)

  return color?.color && color.element
    ? applyPPTXColorModifiers(color.color, color.element)
    : color?.color
}

function readPPTXHexColor(value: string | null | undefined) {
  return value && /^[\da-f]{6}$/i.test(value)
    ? `#${value.toLowerCase()}`
    : undefined
}

function readPPTXSchemeColor(
  value: string | null | undefined,
  themeColors: PPTXThemeColorMap,
) {
  return value ? themeColors[value] : undefined
}

function readPPTXPresetColor(value: string | null | undefined) {
  return value ? PPTX_PRESET_COLORS[value] : undefined
}

function readPPTXSystemColor(value: string | null | undefined) {
  return value ? PPTX_SYSTEM_COLORS[value] : undefined
}

function applyPPTXColorModifiers(color: string, colorElement: Element) {
  let rgb = parsePPTXHexColor(color)

  if (!rgb) {
    return color
  }

  for (const modifier of Array.from(colorElement.children)) {
    const ratio = readPPTXColorModifierRatio(modifier)

    if (ratio === null) {
      continue
    }

    if (modifier.localName === 'shade') {
      rgb = rgb.map((channel) => channel * ratio) as [number, number, number]
    } else if (modifier.localName === 'tint') {
      rgb = rgb.map((channel) =>
        channel + (255 - channel) * ratio) as [number, number, number]
    } else if (modifier.localName === 'lumMod') {
      rgb = rgb.map((channel) => channel * ratio) as [number, number, number]
    } else if (modifier.localName === 'lumOff') {
      rgb = rgb.map((channel) =>
        channel + 255 * ratio) as [number, number, number]
    }
  }

  return formatPPTXHexColor(rgb)
}

function readPPTXColorModifierRatio(modifier: Element) {
  const value = toPPTXPositiveNumber(modifier.getAttribute('val'))

  return value === null
    ? null
    : Math.max(0, Math.min(1, value / 100_000))
}

function parsePPTXHexColor(color: string): [number, number, number] | null {
  const match = color.match(/^#?([\da-f]{6})$/i)

  if (!match) {
    return null
  }

  return [
    Number.parseInt(match[1].slice(0, 2), 16),
    Number.parseInt(match[1].slice(2, 4), 16),
    Number.parseInt(match[1].slice(4, 6), 16),
  ]
}

function formatPPTXHexColor(rgb: readonly number[]) {
  return `#${rgb
    .map((channel) =>
      Math.max(0, Math.min(255, Math.round(channel)))
        .toString(16)
        .padStart(2, '0'))
    .join('')}`
}

function hasPPTXNoFill(container: Element | null) {
  return getDirectPPTXChildByLocalName(container, 'noFill') !== null
}

function readPPTXAlphaOpacity(solidFill: Element) {
  const alpha = getFirstPPTXDescendantByLocalName(solidFill, 'alpha')
  const value = toPPTXPositiveNumber(alpha?.getAttribute('val'))

  return value === null ? null : Math.max(0, Math.min(1, value / 100_000))
}

function readPPTXImageOpacity(blip: Element | null) {
  const alphaModFix = getFirstPPTXDescendantByLocalName(blip, 'alphaModFix')
  const alphaMod = getFirstPPTXDescendantByLocalName(blip, 'alphaMod')
  const alpha = getFirstPPTXDescendantByLocalName(blip, 'alpha')
  const value = toPPTXPositiveNumber(
    alphaModFix?.getAttribute('amt') ??
      alphaMod?.getAttribute('amt') ??
      alpha?.getAttribute('val'),
  )

  if (value === null) {
    return null
  }

  const opacity = Math.max(0, Math.min(1, value / 100_000))

  return opacity === 1 ? null : opacity
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
  const relationshipId = readPPTXRelationshipAttributeId(
    getFirstPPTXDescendantByLocalName(element, 'hlinkClick'),
  )
  const relationship = relationshipId
    ? relationships.get(relationshipId)
    : undefined
  const url = relationship?.targetMode === 'External'
    ? relationship.target
    : undefined

  return url ? { hyperlink: { url } } : null
}

function readPPTXRelationshipAttributeId(element: Element | null) {
  return element?.getAttribute('r:id') ??
    element?.getAttributeNS(PPTX_RELATIONSHIP_ATTRIBUTE_NS, 'id') ??
    element?.getAttribute('id') ??
    null
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

function readPPTXXmlAttribute(attributes: string, name: string) {
  const escapedName = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const localName = name.includes(':') ? name.split(':').at(-1) : name
  const escapedLocalName = localName?.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  const exact = attributes.match(new RegExp(`\\s${escapedName}="([^"]*)"`))
  const byLocalName = escapedLocalName
    ? attributes.match(new RegExp(`\\s(?:[\\w.-]+:)?${escapedLocalName}="([^"]*)"`))
    : null

  return exact?.[1] ?? byLocalName?.[1] ?? null
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

function getPPTXAttributeByLocalName(
  element: Element,
  localName: string,
) {
  return Array.from(element.attributes)
    .find((attribute) => attribute.localName === localName)
    ?.value ?? null
}

function hasPPTXAttribute(
  element: Element | null,
  localName: string,
) {
  return element
    ? Array.from(element.attributes)
      .some((attribute) => attribute.localName === localName)
    : false
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

function isPPTXTrue(value: string | null | undefined) {
  return value === '1' || value === 'true'
}

function clampPPTXPercent(value: number) {
  return Math.max(0, Math.min(100, Math.round(value)))
}

function normalizePPTXAngle(value: number) {
  return ((value % 360) + 360) % 360
}

function emuToPx(value: number) {
  return Math.round(value / PPTX_EMUS_PER_PIXEL)
}

function pointToPx(value: number) {
  return Math.round(value / PPTX_POINTS_PER_PIXEL)
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
