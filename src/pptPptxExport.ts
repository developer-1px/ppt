import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import {
  type PPTComment,
  type PPTDeck,
  type PPTElement,
  type PPTElementAnimation,
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
  type PPTSlideTransition,
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
const PPTX_MARKUP_COMPATIBILITY_NS =
  'http://schemas.openxmlformats.org/markup-compatibility/2006'
const PPTX_POWERPOINT_2010_NS =
  'http://schemas.microsoft.com/office/powerpoint/2010/main'
const PPTX_FLY_IN_MOTION_PATH = 'M 0 0.25 L 0 0 E'

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
type PPTXAnimationTarget = {
  animation: PPTElementAnimation
  objectId: string
}
type PPTXLockDescriptor = {
  lockTagName: string
  lockXml: string
  propertyTagName: string
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
    outputType: 'arraybuffer',
  })
  const arrayBuffer = await toPPTXArrayBuffer(output)

  if (!shouldPatchPPTXPackage(deck)) {
    return createPPTXBlob(arrayBuffer)
  }

  const zip = await JSZip.loadAsync(arrayBuffer)
  await applyPPTXPackagePatches({ deck, zip })
  const patchedOutput = await zip.generateAsync({
    compression: 'DEFLATE',
    mimeType: PPTX_MIME_TYPE,
    type: 'blob',
  })

  return createPPTXBlob(patchedOutput)
}

async function toPPTXArrayBuffer(
  output: Awaited<ReturnType<PPTXPptx['write']>>,
): Promise<ArrayBuffer> {
  if (output instanceof Blob) {
    return output.arrayBuffer()
  }

  if (output instanceof ArrayBuffer) {
    return output
  }

  if (typeof output === 'string') {
    return copyPPTXBytes(new TextEncoder().encode(output))
  }

  if (ArrayBuffer.isView(output)) {
    return copyPPTXBytes(output)
  }

  return new Blob([output as BlobPart]).arrayBuffer()
}

function copyPPTXBytes(view: ArrayBufferView) {
  const bytes = new Uint8Array(view.buffer, view.byteOffset, view.byteLength)
  const copy = new Uint8Array(bytes.byteLength)

  copy.set(bytes)

  return copy.buffer
}

function createPPTXBlob(part: BlobPart | Blob) {
  return part instanceof Blob && part.type === PPTX_MIME_TYPE
    ? part
    : new Blob([part], {
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

function shouldPatchPPTXPackage(deck: PPTDeck) {
  return deck.slides.some((slide) =>
    hasPPTXSlideName(slide) ||
    slide.transition !== undefined ||
    hasPPTXSlideAnimations(slide) ||
    hasPPTXLockedElements(slide) ||
    slide.elements.some((element) =>
      element.visible !== false &&
      Boolean(element.accessibility?.altText.trim())))
}

async function applyPPTXPackagePatches({
  deck,
  zip,
}: {
  deck: PPTDeck
  zip: JSZip
}) {
  await Promise.all(deck.slides.map(async (slide, index) => {
    const path = `ppt/slides/slide${index + 1}.xml`
    const file = zip.file(path)

    if (!file) {
      return
    }

    const xml = await file.async('string')
    const nextXml = setPPTXElementLocksXml(
      setPPTXElementAccessibilityXml(
        setPPTXSlideTimingXml(
          setPPTXSlideTransitionXml(
            setPPTXSlideNameXml(xml, slide),
            createPPTXSlideTransitionXml(slide.transition),
          ),
          slide,
        ),
        slide,
      ),
      slide,
    )

    if (nextXml !== xml) {
      zip.file(path, nextXml)
    }
  }))
}

function hasPPTXSlideName(slide: PPTSlide) {
  return slide.name.trim().length > 0
}

function hasPPTXSlideAnimations(slide: PPTSlide) {
  return slide.elements.some((element) =>
    element.visible !== false &&
    element.animation !== undefined &&
    element.animation.type !== 'none')
}

function hasPPTXLockedElements(slide: PPTSlide) {
  return slide.elements.some((element) =>
    element.visible !== false &&
    element.locked === true)
}

function setPPTXElementAccessibilityXml(xml: string, slide: PPTSlide) {
  return slide.elements.reduce((nextXml, element) => {
    const altText = element.visible === false
      ? ''
      : element.accessibility?.altText.trim() ?? ''

    if (!altText) {
      return nextXml
    }

    return getPPTXElementObjectNames(element).reduce(
      (patchedXml, objectName) =>
        setPPTXObjectDescriptionXml(patchedXml, objectName, altText),
      nextXml,
    )
  }, xml)
}

function setPPTXObjectDescriptionXml(
  xml: string,
  objectName: string,
  description: string,
) {
  const name = escapePPTXXmlAttribute(objectName)
  const descr = escapePPTXXmlAttribute(description)
  const pattern = new RegExp(
    `<p:cNvPr\\b(?=[^>]*\\bname="${escapePPTXRegExp(name)}")[^>]*>`,
    'g',
  )

  return xml.replace(pattern, (tag) => {
    if (tag.includes(' descr=')) {
      return tag.replace(/\sdescr="[^"]*"/, ` descr="${descr}"`)
    }

    return tag.endsWith('/>')
      ? tag.replace(/\/>$/, ` descr="${descr}"/>`)
      : tag.replace(/>$/, ` descr="${descr}">`)
  })
}

function getPPTXElementObjectNames(element: PPTElement) {
  if (element.kind === 'freeform') {
    return element.points.length > 1
      ? element.points
        .slice(1)
        .map((_, index) => `${element.name} segment ${index + 1}`)
      : []
  }

  if (element.kind === 'line' && element.route === 'elbow') {
    return [`${element.name} route`]
  }

  return [element.name]
}

function setPPTXElementLocksXml(xml: string, slide: PPTSlide) {
  return slide.elements.reduce((nextXml, element) => {
    if (element.visible === false || element.locked !== true) {
      return nextXml
    }

    const descriptor = getPPTXLockDescriptor(element)

    return getPPTXElementObjectNames(element).reduce(
      (patchedXml, objectName) =>
        setPPTXObjectLocksXml(patchedXml, objectName, descriptor),
      nextXml,
    )
  }, xml)
}

function getPPTXLockDescriptor(element: PPTElement): PPTXLockDescriptor {
  if (element.kind === 'image') {
    return {
      lockTagName: 'picLocks',
      lockXml: '<a:picLocks noMove="1" noResize="1" noRot="1"/>',
      propertyTagName: 'cNvPicPr',
    }
  }

  if (element.kind === 'line' || element.kind === 'freeform') {
    return {
      lockTagName: 'cxnSpLocks',
      lockXml: '<a:cxnSpLocks noMove="1" noResize="1" noRot="1" noEditPoints="1"/>',
      propertyTagName: 'cNvCxnSpPr',
    }
  }

  if (element.kind === 'table') {
    return {
      lockTagName: 'graphicFrameLocks',
      lockXml: '<a:graphicFrameLocks noMove="1" noResize="1"/>',
      propertyTagName: 'cNvGraphicFramePr',
    }
  }

  return {
    lockTagName: 'spLocks',
    lockXml: '<a:spLocks noMove="1" noResize="1" noRot="1" noTextEdit="1"/>',
    propertyTagName: 'cNvSpPr',
  }
}

function setPPTXObjectLocksXml(
  xml: string,
  objectName: string,
  descriptor: PPTXLockDescriptor,
) {
  const name = escapePPTXXmlAttribute(objectName)
  const segmentPattern = new RegExp(
    `<p:cNvPr\\b(?=[^>]*\\bname="${escapePPTXRegExp(name)}")[\\s\\S]*?</p:nv(?:Sp|Pic|CxnSp|GraphicFrame)Pr>`,
    'g',
  )

  return xml.replace(segmentPattern, (segment) =>
    setPPTXNonVisualPropertyLocksXml(segment, descriptor))
}

function setPPTXNonVisualPropertyLocksXml(
  segment: string,
  descriptor: PPTXLockDescriptor,
) {
  const lockPattern = new RegExp(`<a:${descriptor.lockTagName}\\b[^>]*/>`)

  if (lockPattern.test(segment)) {
    return segment.replace(lockPattern, (tag) =>
      mergePPTXLockXmlAttributes(tag, descriptor.lockXml))
  }

  const propertyPattern = new RegExp(
    `<p:${descriptor.propertyTagName}\\b([^>]*)/>`,
  )

  if (propertyPattern.test(segment)) {
    return segment.replace(
      propertyPattern,
      `<p:${descriptor.propertyTagName}$1>${descriptor.lockXml}</p:${descriptor.propertyTagName}>`,
    )
  }

  return segment.replace(
    new RegExp(`</p:${descriptor.propertyTagName}>`),
    `${descriptor.lockXml}</p:${descriptor.propertyTagName}>`,
  )
}

function mergePPTXLockXmlAttributes(tag: string, lockXml: string) {
  return getPPTXXmlAttributeEntries(lockXml).reduce(
    (nextTag, [name, value]) => setPPTXXmlTagAttribute(nextTag, name, value),
    tag,
  )
}

function getPPTXXmlAttributeEntries(tag: string) {
  return [...tag.matchAll(/\s([A-Za-z_:][A-Za-z0-9_.:-]*)="([^"]*)"/g)]
    .map((match) => [match[1], match[2]] as const)
}

function setPPTXXmlTagAttribute(tag: string, name: string, value: string) {
  const escapedName = escapePPTXRegExp(name)
  const escapedValue = escapePPTXXmlAttribute(value)
  const attributePattern = new RegExp(`\\s${escapedName}="[^"]*"`)

  if (attributePattern.test(tag)) {
    return tag.replace(attributePattern, ` ${name}="${escapedValue}"`)
  }

  return tag.endsWith('/>')
    ? tag.replace(/\/>$/, ` ${name}="${escapedValue}"/>`)
    : tag.replace(/>$/, ` ${name}="${escapedValue}">`)
}

function setPPTXSlideNameXml(xml: string, slide: PPTSlide) {
  const name = slide.name.trim()

  if (!name) {
    return xml
  }

  return xml.replace(/<p:cSld\b[^>]*>/, (tag) =>
    setPPTXXmlTagAttribute(tag, 'name', name))
}

function setPPTXSlideTimingXml(xml: string, slide: PPTSlide) {
  if (!hasPPTXSlideAnimations(slide)) {
    return xml
  }

  const xmlWithoutTiming = xml.replace(
    /<p:timing\b[\s\S]*?<\/p:timing>/,
    '',
  )
  const timingXml = createPPTXSlideTimingXml(xmlWithoutTiming, slide)

  if (!timingXml) {
    return xml
  }

  if (xmlWithoutTiming.includes('</p:transition>')) {
    return xmlWithoutTiming.replace(
      '</p:transition>',
      `</p:transition>${timingXml}`,
    )
  }

  if (/<p:transition\b[^/]*\/>/.test(xmlWithoutTiming)) {
    return xmlWithoutTiming.replace(
      /<p:transition\b[^/]*\/>/,
      (tag) => `${tag}${timingXml}`,
    )
  }

  const anchor = xmlWithoutTiming.includes('</p:clrMapOvr>')
    ? '</p:clrMapOvr>'
    : '</p:cSld>'

  return xmlWithoutTiming.replace(anchor, `${anchor}${timingXml}`)
}

function createPPTXSlideTimingXml(xml: string, slide: PPTSlide) {
  const targets = getPPTXSlideAnimationTargets(xml, slide)

  if (targets.length === 0) {
    return null
  }

  let nextTimeNodeId = 3
  const effectXml = targets.map((target) => {
    const startId = nextTimeNodeId
    nextTimeNodeId += 3

    return createPPTXAnimationEffectXml(target, startId)
  }).join('')
  const buildXml = targets.map(({ objectId }) =>
    `<p:bldP spid="${objectId}" grpId="0" build="allAtOnce"/>`).join('')

  return [
    '<p:timing>',
    '<p:tnLst>',
    '<p:par>',
    '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">',
    '<p:childTnLst>',
    '<p:seq concurrent="1" nextAc="seek">',
    '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">',
    `<p:childTnLst>${effectXml}</p:childTnLst>`,
    '</p:cTn>',
    '</p:seq>',
    '</p:childTnLst>',
    '</p:cTn>',
    '</p:par>',
    '</p:tnLst>',
    `<p:bldLst>${buildXml}</p:bldLst>`,
    '</p:timing>',
  ].join('')
}

function getPPTXSlideAnimationTargets(xml: string, slide: PPTSlide) {
  const targets: PPTXAnimationTarget[] = []
  const animatedElements = slide.elements
    .filter((element) =>
      element.visible !== false &&
      element.animation !== undefined &&
      element.animation.type !== 'none')
    .sort((a, b) => (a.animation?.order ?? 0) - (b.animation?.order ?? 0))

  for (const element of animatedElements) {
    const objectIds = new Set(
      getPPTXElementObjectNames(element)
        .flatMap((objectName) => getPPTXObjectIdsByName(xml, objectName)),
    )

    for (const objectId of objectIds) {
      if (element.animation) {
        targets.push({
          animation: element.animation,
          objectId,
        })
      }
    }
  }

  return targets
}

function getPPTXObjectIdsByName(xml: string, objectName: string) {
  const name = escapePPTXXmlAttribute(objectName)
  const objectIds: string[] = []

  for (const match of xml.matchAll(/<p:cNvPr\b[^>]*>/g)) {
    const tag = match[0]

    if (getPPTXXmlTagAttribute(tag, 'name') === name) {
      const objectId = getPPTXXmlTagAttribute(tag, 'id')

      if (objectId) {
        objectIds.push(objectId)
      }
    }
  }

  return objectIds
}

function getPPTXXmlTagAttribute(tag: string, name: string) {
  const pattern = new RegExp(`\\b${escapePPTXRegExp(name)}="([^"]*)"`)

  return tag.match(pattern)?.[1] ?? null
}

function createPPTXAnimationEffectXml(
  target: PPTXAnimationTarget,
  timeNodeId: number,
) {
  const presetId = target.animation.type === 'flyIn' ? '2' : '10'
  const presetSubtype = target.animation.type === 'flyIn' ? '8' : '0'
  const nodeType = target.animation.trigger === 'withPrevious'
    ? 'withEffect'
    : 'clickEffect'
  const startDelay = target.animation.trigger === 'withPrevious'
    ? clampPPTXAnimationMs(target.animation.delayMs)
    : 'indefinite'
  const effectDelay = target.animation.trigger === 'withPrevious'
    ? 0
    : clampPPTXAnimationMs(target.animation.delayMs)
  const effectXml = target.animation.type === 'flyIn'
    ? createPPTXAnimationMotionXml(target, timeNodeId + 2, effectDelay)
    : createPPTXAnimationFadeXml(target, timeNodeId + 2, effectDelay)

  return [
    '<p:par>',
    `<p:cTn id="${timeNodeId}" presetID="${presetId}" presetClass="entr" `,
    `presetSubtype="${presetSubtype}" fill="hold" nodeType="${nodeType}">`,
    '<p:stCondLst>',
    `<p:cond delay="${startDelay}"/>`,
    '</p:stCondLst>',
    '<p:childTnLst>',
    createPPTXAnimationVisibilityXml(target, timeNodeId + 1),
    effectXml,
    '</p:childTnLst>',
    '</p:cTn>',
    '</p:par>',
  ].join('')
}

function createPPTXAnimationVisibilityXml(
  target: PPTXAnimationTarget,
  timeNodeId: number,
) {
  return [
    '<p:set>',
    '<p:cBhvr>',
    `<p:cTn id="${timeNodeId}" dur="1" fill="hold"/>`,
    createPPTXAnimationTargetXml(target.objectId),
    '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>',
    '</p:cBhvr>',
    '<p:to><p:strVal val="visible"/></p:to>',
    '</p:set>',
  ].join('')
}

function createPPTXAnimationFadeXml(
  target: PPTXAnimationTarget,
  timeNodeId: number,
  delayMs: number,
) {
  return [
    '<p:animEffect transition="in" filter="fade">',
    createPPTXAnimationBehaviorXml(target, timeNodeId, delayMs),
    '</p:animEffect>',
  ].join('')
}

function createPPTXAnimationMotionXml(
  target: PPTXAnimationTarget,
  timeNodeId: number,
  delayMs: number,
) {
  return [
    `<p:animMotion origin="layout" path="${PPTX_FLY_IN_MOTION_PATH}" `,
    'pathEditMode="relative">',
    createPPTXAnimationBehaviorXml(target, timeNodeId, delayMs),
    '</p:animMotion>',
  ].join('')
}

function createPPTXAnimationBehaviorXml(
  target: PPTXAnimationTarget,
  timeNodeId: number,
  delayMs: number,
) {
  return [
    '<p:cBhvr>',
    `<p:cTn id="${timeNodeId}" dur="${clampPPTXAnimationMs(target.animation.durationMs)}" fill="hold">`,
    '<p:stCondLst>',
    `<p:cond delay="${delayMs}"/>`,
    '</p:stCondLst>',
    '</p:cTn>',
    createPPTXAnimationTargetXml(target.objectId),
    '</p:cBhvr>',
  ].join('')
}

function createPPTXAnimationTargetXml(objectId: string) {
  return `<p:tgtEl><p:spTgt spid="${escapePPTXXmlAttribute(objectId)}"/></p:tgtEl>`
}

function setPPTXSlideTransitionXml(
  xml: string,
  transitionXml: string | null,
) {
  const xmlWithoutTransition = xml.replace(
    /<p:transition\b[\s\S]*?<\/p:transition>|<p:transition\b[^/]*\/>/,
    '',
  )

  if (!transitionXml) {
    return xmlWithoutTransition
  }

  const xmlWithNamespaces = ensurePPTXSlideTransitionNamespaces(
    xmlWithoutTransition,
  )
  const anchor = xmlWithNamespaces.includes('</p:clrMapOvr>')
    ? '</p:clrMapOvr>'
    : '</p:cSld>'

  return xmlWithNamespaces.replace(anchor, `${anchor}${transitionXml}`)
}

function ensurePPTXSlideTransitionNamespaces(xml: string) {
  return ensurePPTXIgnorableNamespace(
    ensurePPTXRootNamespace(
      ensurePPTXRootNamespace(xml, 'p14', PPTX_POWERPOINT_2010_NS),
      'mc',
      PPTX_MARKUP_COMPATIBILITY_NS,
    ),
    'p14',
  )
}

function ensurePPTXRootNamespace(
  xml: string,
  prefix: string,
  namespace: string,
) {
  if (xml.includes(`xmlns:${prefix}=`)) {
    return xml
  }

  return xml.replace('<p:sld ', `<p:sld xmlns:${prefix}="${namespace}" `)
}

function ensurePPTXIgnorableNamespace(xml: string, prefix: string) {
  return xml.replace(/<p:sld\b([^>]*)>/, (tag, attrs: string) => {
    const ignorable = attrs.match(/\smc:Ignorable="([^"]*)"/)

    if (!ignorable) {
      return tag.replace('<p:sld', `<p:sld mc:Ignorable="${prefix}"`)
    }

    const prefixes = ignorable[1].split(/\s+/).filter(Boolean)

    if (prefixes.includes(prefix)) {
      return tag
    }

    return tag.replace(
      /\smc:Ignorable="[^"]*"/,
      ` mc:Ignorable="${[...prefixes, prefix].join(' ')}"`,
    )
  })
}

function createPPTXSlideTransitionXml(
  transition: PPTSlideTransition | undefined,
) {
  if (!transition) {
    return null
  }

  const attributes = [
    `advClick="${transition.advanceOnClick === false ? '0' : '1'}"`,
    `p14:dur="${clampPPTXTransitionMs(transition.durationMs)}"`,
    `spd="${getPPTXSlideTransitionSpeed(transition.durationMs)}"`,
  ]

  if (transition.advanceAfterMs !== null &&
    transition.advanceAfterMs !== undefined) {
    attributes.push(`advTm="${clampPPTXTransitionMs(transition.advanceAfterMs)}"`)
  }

  const childXml = getPPTXSlideTransitionChildXml(transition.type)

  return childXml
    ? `<p:transition ${attributes.join(' ')}>${childXml}</p:transition>`
    : `<p:transition ${attributes.join(' ')}/>`
}

function getPPTXSlideTransitionChildXml(
  type: PPTSlideTransition['type'],
) {
  if (type === 'fade') {
    return '<p:fade/>'
  }

  if (type === 'push') {
    return '<p:push dir="l"/>'
  }

  return null
}

function getPPTXSlideTransitionSpeed(durationMs: number) {
  if (durationMs <= 500) {
    return 'fast'
  }

  return durationMs <= 1000 ? 'med' : 'slow'
}

function clampPPTXTransitionMs(value: number) {
  return Math.round(clamp(value, 0, 2_147_483_647))
}

function clampPPTXAnimationMs(value: number) {
  return Math.round(clamp(value, 0, 2_147_483_647))
}

function escapePPTXXmlAttribute(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}

function escapePPTXRegExp(value: string) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
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
  const hyperlink = createPPTXHyperlink(element.hyperlink)
  const opacity = getPPTXElementOpacity(element)

  pptxSlide.addText(createPPTXTextRuns({
    body: element.textBody,
    hyperlink,
    opacity,
    style: element.style,
  }), {
    ...createPPTXElementTextOptions({
      element,
      hyperlink,
      inset: getPPTXTextInset(element.style, PPTX_DEFAULT_TEXT_BOX_INSET),
      opacity,
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
  const opacity = getPPTXElementOpacity(element)
  const shapeOptions = {
    ...createPPTXElementPosition(element.geometry),
    fill: createPPTXFill(element.fill, opacity),
    flipH: element.flipH === true,
    flipV: element.flipV === true,
    hyperlink: createPPTXHyperlink(element.hyperlink),
    line: createPPTXLineProps(element.stroke, {}, opacity),
    objectName: element.name,
    rectRadius: getPPTXRectRadius(element),
    shadow: createPPTXShadow(element.shadow, opacity),
  }

  if (element.textBody) {
    const hyperlink = createPPTXHyperlink(element.hyperlink)

    pptxSlide.addText(createPPTXTextRuns({
      body: element.textBody,
      hyperlink,
      opacity,
      style: element.style,
    }), {
      ...shapeOptions,
      ...createPPTXTextStyleOptions(element.style, opacity),
      fit: getPPTXTextFit(element.textAutoFit),
      hyperlink,
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
  const opacity = getPPTXElementOpacity(element)

  pptxSlide.addImage({
    ...source,
    ...position,
    altText: element.accessibility?.altText ?? element.alt,
    flipH: element.flipH === true,
    flipV: element.flipV === true,
    hyperlink: createPPTXHyperlink(element.hyperlink),
    objectName: element.name,
    rotate: element.geometry.rotation,
    shadow: createPPTXShadow(element.shadow, opacity),
    sizing: createPPTXImageSizing(element, position),
    transparency: toPPTXTransparency(opacity),
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
  const opacity = getPPTXElementOpacity(element)

  if (element.route === 'elbow') {
    addPPTXElbowRoute({
      element,
      pptxSlide,
      opacity,
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
    }, opacity),
    objectName: element.name,
    rotate: element.geometry.rotation,
    shadow: createPPTXShadow(element.shadow, opacity),
    w: pxToIn(end.x - start.x),
    x: pxToIn(start.x),
    y: pxToIn(start.y),
  })

}

function addPPTXElbowRoute({
  element,
  end,
  opacity,
  pptxSlide,
  start,
}: {
  element: PPTLine
  end: { x: number, y: number }
  opacity: number
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
      }, opacity),
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
  const opacity = getPPTXElementOpacity(element)

  for (let index = 1; index < element.points.length; index += 1) {
    const start = element.points[index - 1]
    const end = element.points[index]

    pptxSlide.addShape('line', {
      h: pxToIn(end.y - start.y),
      line: createPPTXLineProps(element.stroke, {}, opacity),
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
  const opacity = getPPTXElementOpacity(element)
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
        fill: {
          color: rowIndex === 0 ? 'EFF6FF' : 'FFFFFF',
          transparency: toPPTXTransparency(opacity),
        },
        fontFace: PPTX_DEFAULT_FONT_FACE,
        fontSize: 13.5,
        margin: 0.08,
        transparency: toPPTXTransparency(opacity),
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
  const opacity = getPPTXElementOpacity(element)
  const body = [
    element.body,
    ...(element.thread ?? []).map((message) =>
      `${message.authorName}: ${message.body}`),
  ].filter(Boolean).join('\n')

  pptxSlide.addText(body, {
    ...createPPTXElementPosition(element.geometry),
    color: '78350F',
    fill: {
      color: element.resolved ? 'FEF3C7' : 'FFFBEB',
      transparency: toPPTXTransparency(opacity),
    },
    fontFace: PPTX_DEFAULT_FONT_FACE,
    fontSize: 12.75,
    line: createPPTXLineProps({ color: '#D97706', width: 1 }, {}, opacity),
    margin: [6, 8, 6, 8],
    objectName: element.name,
    shape: 'roundRect',
    transparency: toPPTXTransparency(opacity),
  })
}

function createPPTXElementTextOptions({
  element,
  hyperlink,
  inset,
  opacity,
}: {
  element: Extract<PPTElement, { kind: 'textBox' }>
  hyperlink: PptxGenJS.HyperlinkProps | undefined
  inset: PPTXTextInset
  opacity: number
}) {
  return {
    ...createPPTXElementPosition(element.geometry),
    ...createPPTXTextStyleOptions(element.style, opacity),
    fit: getPPTXTextFit(element.textAutoFit),
    hyperlink,
    margin: createPPTXMargin(inset),
    objectName: element.name,
    shadow: createPPTXShadow(element.shadow, opacity),
    valign: element.style.verticalAlign ?? 'top',
  }
}

function createPPTXTextRuns({
  body,
  hyperlink,
  opacity = 1,
  style,
}: {
  body: PPTTextBody
  hyperlink?: PptxGenJS.HyperlinkProps
  opacity?: number
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
          ...createPPTXTextStyleOptions(style, opacity),
          ...createPPTXParagraphOptions(paragraph),
          ...createPPTXTextRunOptions(run),
          breakLine: paragraphIndex > 0 && runIndex === 0,
          hyperlink,
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
  opacity = 1,
): Partial<PptxGenJS.TextPropsOptions> {
  return {
    bold: style?.fontWeight === 'bold',
    color: toPPTXColor(style?.color ?? PPTX_DEFAULT_TEXT_COLOR, '111827'),
    fontFace: style?.fontFamily ?? PPTX_DEFAULT_FONT_FACE,
    fontSize: pxToPt(style?.fontSize ?? PPTX_DEFAULT_FONT_SIZE),
    transparency: toPPTXTransparency(opacity),
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

function createPPTXFill(fill: PPTFill, opacity = 1) {
  return {
    color: toPPTXColor(fill.color, 'FFFFFF'),
    transparency: toPPTXTransparency((fill.opacity ?? 1) * opacity),
  }
}

function createPPTXLineProps(
  stroke: PPTStroke | undefined,
  markers: {
    endMarker?: PPTLine['endMarker']
    startMarker?: PPTLine['startMarker']
  } = {},
  opacity = 1,
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
    transparency: toPPTXTransparency(opacity),
    width: Math.max(0.25, pxToPt(stroke.width)),
  }
}

function createPPTXHyperlink(
  hyperlink: PPTElement['hyperlink'] | undefined,
): PptxGenJS.HyperlinkProps | undefined {
  return hyperlink?.url
    ? { url: hyperlink.url }
    : undefined
}

function createPPTXImageSizing(
  element: PPTImage,
  position: PPTXPosition,
): PptxGenJS.ImageProps['sizing'] {
  if (element.crop &&
    (element.crop.x !== 50 || element.crop.y !== 50)) {
    return {
      h: position.h,
      type: 'crop',
      w: position.w,
      x: getPPTXImageCropOffset(element.crop.x, position.w),
      y: getPPTXImageCropOffset(element.crop.y, position.h),
    }
  }

  return {
    h: position.h,
    type: element.fit === 'contain' ? 'contain' : 'cover',
    w: position.w,
  }
}

function getPPTXImageCropOffset(value: number, size: number) {
  return ((clamp(value, 0, 100) - 50) / 100) * size
}

function getPPTXElementOpacity(element: Pick<PPTElement, 'opacity'>) {
  return clamp(element.opacity ?? 1, 0, 1)
}

function createPPTXShadow(
  shadow: PPTElementShadow | undefined,
  opacity = 1,
) {
  if (!shadow) {
    return undefined
  }

  return {
    angle: shadow.angle,
    blur: pxToPt(shadow.blur),
    color: toPPTXColor(shadow.color, '000000'),
    offset: pxToPt(shadow.distance),
    opacity: clamp(shadow.opacity * opacity, 0, 1),
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
