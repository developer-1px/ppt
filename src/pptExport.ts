import {
  getSlideEditColorWithAlphaCSS,
  getSlideEditObjectShadowFilter,
  getSlideEditObjectStrokeLineStyleBorderStyle,
  getSlideEditObjectStrokeLineStyleDashArray,
  getSlideEditTextParagraphListLevelIndentCSSValue,
  getSlideEditTextParagraphListLevelIndentEm,
  normalizeSlideEditTextParagraphListLevel,
  getSlideEditTextFrameInsetPaddingCSS,
  getSlideEditTextVerticalAlignmentFlexAlignItems,
  normalizeSlideEditObjectCornerRadius,
  normalizeSlideEditObjectFillOpacity,
  normalizeSlideEditObjectOpacity,
  normalizeSlideEditObjectStrokeLineStyle,
  normalizeSlideEditTextFrameInsetValue,
  normalizeSlideEditTextVerticalAlignment,
  toSlideEditObjectCornerRadiusAttributeValue,
  toSlideEditObjectFillOpacityAttributeValue,
  toSlideEditObjectOpacityAttributeValue,
} from './pptSlideEditAffordanceAdapter'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTComment,
  type PPTCommentThreadMessage,
  type PPTDeck,
  type PPTElement,
  type PPTElementAccessibility,
  type PPTElementHyperlink,
  type PPTElementShadow,
  type PPTFill,
  type PPTFreeform,
  type PPTImage,
  type PPTLine,
  type PPTLinePoint,
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTSlide,
  type PPTStroke,
  type PPTStrokeDash,
  type PPTTable,
  type PPTTableCellBorders,
  type PPTTableCellTextStyle,
  type PPTTextBody,
  type PPTTextStyle,
} from './pptModel'
import {
  getPPTTableCellBorders,
  getPPTTableCellColSpan,
  getPPTTableCellFill,
  getPPTTableCellRowSpan,
  getPPTTableCellTextStyle,
  getPPTTableResolvedColumnWidths,
  getPPTTableResolvedRowHeights,
  isPPTTableCellHidden,
} from './pptTableLayout'
import {
  createPPTCanvasCssBoundsTransform,
  createPPTCanvasSvgBoundsTransform,
  createPPTCanvasSvgFreehandPathData,
  createPPTCanvasSvgPathData,
  createPPTCanvasSvgPathSegmentData,
  escapePPTCanvasXmlAttribute,
  formatPPTCanvasSvgNumber,
} from './pptCanvasRendererAdapter'
import { unionPPTCanvasRectList } from './pptCanvasFoundationAdapter'

const PPT_SELECTION_EXPORT_PADDING = 24
const PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT = 1.14
const PPT_PARAGRAPH_LINE_HEIGHT_MIN = 0.8
const PPT_PARAGRAPH_LINE_HEIGHT_MAX = 3
const PPT_PARAGRAPH_SPACING_MAX = 240
const PPT_DEFAULT_TEXT_FONT_FAMILY = 'Inter'
const PPT_TEXT_FONT_FAMILY_OPTIONS = Object.freeze([
  { css: 'Inter, ui-sans-serif, system-ui, sans-serif', value: 'Inter' },
  { css: 'Arial, Helvetica, sans-serif', value: 'Arial' },
  { css: 'Georgia, serif', value: 'Georgia' },
  { css: '"Courier New", monospace', value: 'Courier New' },
] as const)
const PPT_TEXT_FONT_FAMILY_VALUES = new Set<string>(
  PPT_TEXT_FONT_FAMILY_OPTIONS.map((option) => option.value),
)
type PPTTextVerticalAlign = NonNullable<PPTTextStyle['verticalAlign']>
const PPT_DEFAULT_TEXT_VERTICAL_ALIGN: PPTTextVerticalAlign = 'top'
type PPTTextInset = NonNullable<PPTTextStyle['textInset']>
const PPT_DEFAULT_TEXT_BOX_INSET = Object.freeze({
  bottom: 0,
  left: 0,
  right: 0,
  top: 0,
} as const satisfies PPTTextInset)
const PPT_DEFAULT_SHAPE_TEXT_INSET = Object.freeze({
  bottom: 18,
  left: 18,
  right: 18,
  top: 18,
} as const satisfies PPTTextInset)
const PPT_DEFAULT_ELEMENT_SHADOW = Object.freeze({
  angle: 45,
  blur: 14,
  color: '#000000',
  distance: 8,
  opacity: 0.22,
} as const satisfies PPTElementShadow)
const PPT_ELEMENT_SHADOW_ANGLE_MIN = 0
const PPT_ELEMENT_SHADOW_ANGLE_MAX = 359
const PPT_ELEMENT_SHADOW_BLUR_MAX = 80
const PPT_ELEMENT_SHADOW_DISTANCE_MAX = 120
const PPT_ELEMENT_SHADOW_OPACITY_MIN = 0
const PPT_ELEMENT_SHADOW_OPACITY_MAX = 1
const PPT_ALT_TEXT_MAX_LENGTH = 1000
const PPT_SHAPE_CORNER_RADIUS_DEFAULT = 24
const PPT_HYPERLINK_URL_MAX_LENGTH = 2048

export function exportPPTDeckHTML(deck: PPTDeck) {
  const body = deck.slides.map((slide) => {
    const elements = slide.elements
      .filter((element) => element.visible !== false)
      .map(renderPPTElementHTML)
      .join('\n')

    const layoutAttr = slide.layoutId
      ? ` data-ppt-layout-id="${escapeHtml(slide.layoutId)}"`
      : ''
    const themeAttr = slide.themeId
      ? ` data-ppt-theme-id="${escapeHtml(slide.themeId)}"`
      : ''
    const hiddenAttr = slide.hidden === true
      ? ' data-ppt-slide-hidden="true"'
      : ''
    const sectionNameAttr = slide.sectionName
      ? ` data-ppt-slide-section-name="${escapeHtml(slide.sectionName)}"`
      : ''
    const placeholderVisibilityAttr = getPPTSlideHiddenPlaceholderAttr(
      slide,
      'data-ppt-hidden-placeholders',
    )
    const transitionAttrs = getPPTSlideTransitionAttrs(slide, 'data-ppt-transition')

    return [
      `  <section class="ppt-slide" data-ppt-slide="${escapeHtml(slide.id)}"${layoutAttr}${themeAttr}${hiddenAttr}${sectionNameAttr}${placeholderVisibilityAttr}${transitionAttrs} style="background:${escapeHtml(slide.background ? getPPTFillColorCSS(slide.background) : '#ffffff')}">`,
      elements,
      '  </section>',
      renderPPTSlideNotesHTML(slide.id, slide.notes),
    ].join('\n')
  }).join('\n')

  return [
    '<!doctype html>',
    '<html lang="en">',
    '<head>',
    '  <meta charset="utf-8" />',
    '  <meta name="viewport" content="width=device-width, initial-scale=1" />',
    `  <title>${escapeHtml(deck.title)}</title>`,
    '  <style>',
    indent(exportCSS(), 4),
    '  </style>',
    '</head>',
    '<body>',
    '<main class="ppt-deck">',
    body,
    '</main>',
    '<script type="application/json" data-ppt-deck>',
    indent(escapeScriptJson(JSON.stringify(deck, null, 2)), 2),
    '</script>',
    '</body>',
    '</html>',
    '',
  ].join('\n')
}

export function exportPPTSlideSVG(slide: PPTSlide) {
  return renderPPTElementsSVG({
    elements: getVisiblePPTElements(slide.elements),
    height: PPT_SLIDE_HEIGHT,
    scope: 'slide',
    slide,
    viewBox: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    width: PPT_SLIDE_WIDTH,
  })
}

export function exportPPTSelectionSVG(
  slide: PPTSlide,
  selection: readonly string[],
) {
  if (selection.length === 0) {
    return null
  }

  const selected = new Set(selection)
  const elements = getVisiblePPTElements(slide.elements)
    .filter((element) => selected.has(element.id))
  const bounds = getPPTElementsExportBounds(elements)

  if (!bounds) {
    return null
  }

  const viewBox = {
    h: Math.max(1, bounds.h + PPT_SELECTION_EXPORT_PADDING * 2),
    w: Math.max(1, bounds.w + PPT_SELECTION_EXPORT_PADDING * 2),
    x: bounds.x - PPT_SELECTION_EXPORT_PADDING,
    y: bounds.y - PPT_SELECTION_EXPORT_PADDING,
  }

  return renderPPTElementsSVG({
    elements,
    height: Math.ceil(viewBox.h),
    scope: 'selection',
    slide,
    viewBox,
    width: Math.ceil(viewBox.w),
  })
}

function renderPPTElementsSVG({
  elements,
  height,
  scope,
  slide,
  viewBox,
  width,
}: {
  elements: PPTElement[]
  height: number
  scope: 'selection' | 'slide'
  slide: PPTSlide
  viewBox: PPTElement['geometry']
  width: number
}) {
  const background = slide.background?.color ?? '#ffffff'

  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="${formatNumber(viewBox.x)} ${formatNumber(viewBox.y)} ${formatNumber(viewBox.w)} ${formatNumber(viewBox.h)}" data-ppt-svg-slide="${escapeHtml(slide.id)}" data-ppt-svg-scope="${scope}"${getPPTSlideLayoutThemeSvgAttrs(slide)}>`,
    `<rect data-ppt-svg-background="true" x="${formatNumber(viewBox.x)}" y="${formatNumber(viewBox.y)}" width="${formatNumber(viewBox.w)}" height="${formatNumber(viewBox.h)}" fill="${escapeHtml(background)}" />`,
    ...elements.map(renderPPTElementSVG),
    '</svg>',
    '',
  ].join('\n')
}

function getPPTSlideLayoutThemeSvgAttrs(slide: PPTSlide) {
  const attrs = [
    slide.layoutId ? `data-ppt-svg-layout-id="${escapeHtml(slide.layoutId)}"` : '',
    slide.themeId ? `data-ppt-svg-theme-id="${escapeHtml(slide.themeId)}"` : '',
    getPPTSlideHiddenPlaceholderAttr(slide, 'data-ppt-svg-hidden-placeholders').trim(),
    getPPTSlideTransitionAttrs(slide, 'data-ppt-svg-transition').trim(),
  ].filter(Boolean)

  return attrs.length > 0 ? ` ${attrs.join(' ')}` : ''
}

function getPPTSlideHiddenPlaceholderAttr(slide: PPTSlide, attributeName: string) {
  const ids = (slide.hiddenPlaceholderIds ?? []).filter(Boolean)

  return ids.length > 0
    ? ` ${attributeName}="${escapeHtml(ids.join(' '))}"`
    : ''
}

function getPPTSlideTransitionAttrs(slide: PPTSlide, prefix: string) {
  if (!slide.transition) {
    return ''
  }

  const attrs = [
    `${prefix}-type="${escapeHtml(slide.transition.type)}"`,
    `${prefix}-duration="${slide.transition.durationMs}"`,
    `${prefix}-advance-on-click="${slide.transition.advanceOnClick === false ? 'false' : 'true'}"`,
  ]

  if (slide.transition.advanceAfterMs !== null && slide.transition.advanceAfterMs !== undefined) {
    attrs.push(`${prefix}-advance-after="${slide.transition.advanceAfterMs}"`)
  }

  return ` ${attrs.join(' ')}`
}

function renderPPTElementHTML(element: PPTElement) {
  const transform = getPPTElementTransform(element)
  const transformAttrs = getPPTElementTransformAttrs(element)
  const altTextAttr = getPPTElementAltTextHTMLAttr(element)
  const hyperlinkAttr = getPPTElementHyperlinkHTMLAttr(element)
  const cornerRadiusAttr = getPPTElementCornerRadiusHTMLAttr(element)
  const fillOpacityAttr = getPPTElementFillOpacityHTMLAttr(element)
  const strokeDashAttr = getPPTElementStrokeDashHTMLAttr(element)
  const opacity = getPPTElementOpacity(element)
  const opacityAttr = ` data-ppt-opacity="${escapeHtml(formatPPTElementOpacity(opacity))}"`
  const shadowAttrs = getPPTElementShadowHTMLAttrs(element)
  const style = [
    `left:${toPercent(element.geometry.x, PPT_SLIDE_WIDTH)}`,
    `top:${toPercent(element.geometry.y, PPT_SLIDE_HEIGHT)}`,
    `width:${toPercent(element.geometry.w, PPT_SLIDE_WIDTH)}`,
    `height:${toPercent(element.geometry.h, PPT_SLIDE_HEIGHT)}`,
    `opacity:${formatPPTElementOpacity(opacity)}`,
    getPPTElementFilterStyle(element),
    transform
      ? `transform:${transform}`
      : '',
    transform
      ? 'transform-origin:center'
      : '',
  ]

  if (element.kind === 'image') {
    const crop = element.crop ?? { x: 50, y: 50 }
    const fit = element.fit ?? 'cover'

    return `    <img class="ppt-element ppt-image" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${altTextAttr}${hyperlinkAttr}${strokeDashAttr}${opacityAttr}${shadowAttrs} data-ppt-image-fit="${fit}" data-ppt-image-crop-x="${crop.x}" data-ppt-image-crop-y="${crop.y}"${getPPTImageCropRectHTMLAttrs(element)}${getPPTImageAdjustmentsHTMLAttrs(element)}${getPPTImageClipShapeHTMLAttr(element)} alt="${escapeHtml(getPPTImageAltText(element))}" src="${escapeHtml(element.src)}" style="${[...style, `object-fit:${fit}`, `object-position:${crop.x}% ${crop.y}%`, getPPTImageClipShapeStyle(element), getPPTImageStrokeStyle(element)].filter(Boolean).join(';')}" />`
  }

  if (element.kind === 'line') {
    return renderPPTLineHTML(element, style)
  }

  if (element.kind === 'freeform') {
    return renderPPTFreeformHTML(element, style)
  }

  if (element.kind === 'table') {
    return renderPPTTableHTML(element, style, transformAttrs)
  }

  if (element.kind === 'comment') {
    return renderPPTCommentHTML(element, style, transformAttrs)
  }

  const text = renderPPTTextBodyHTML(element.textBody)
  const textStyle = element.style ? exportTextStyle(element.style) : ''
  const paragraphStyle = `text-align:${element.textBody?.paragraphs[0]?.align ?? 'left'}`
  const textInset = getPPTElementTextInset(element)
  const textInsetAttr = ` data-ppt-text-inset="${escapeHtml(formatPPTTextInsetData(textInset))}"`
  const textInsetStyle = `padding:${getPPTTextInsetCSS(textInset)}`
  const verticalAlign = getPPTElementTextVerticalAlign(element)
  const verticalAlignAttr = ` data-ppt-vertical-align="${escapeHtml(verticalAlign)}"`
  const verticalAlignStyle = `align-items:${getPPTTextVerticalAlignCSS(verticalAlign)}`
  const fontFamilyAttr = element.style
    ? ` data-ppt-font-family="${escapeHtml(normalizePPTTextFontFamily(element.style.fontFamily))}"`
    : ''
  const bulletListAttr = element.textBody && hasPPTTextBodyBullet(element.textBody)
    ? ' data-ppt-bullet-list="true"'
    : ''
  const numberedListAttr = element.textBody && hasPPTTextBodyNumbered(element.textBody)
    ? ' data-ppt-numbered-list="true"'
    : ''
  const autoFitAttr = getPPTTextAutoFitAttr(element)

  if (element.kind === 'shape') {
    return `    <div class="ppt-element ppt-shape ppt-shape-${element.shape}" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${altTextAttr}${hyperlinkAttr}${cornerRadiusAttr}${fillOpacityAttr}${strokeDashAttr}${opacityAttr}${shadowAttrs}${fontFamilyAttr}${textInsetAttr}${verticalAlignAttr}${bulletListAttr}${numberedListAttr}${autoFitAttr} style="${[...style, exportShapeStyle(element), textStyle, textInsetStyle, verticalAlignStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
  }

  return `    <div class="ppt-element ppt-text" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${altTextAttr}${hyperlinkAttr}${opacityAttr}${shadowAttrs}${fontFamilyAttr}${textInsetAttr}${verticalAlignAttr}${bulletListAttr}${numberedListAttr}${autoFitAttr} style="${[...style, textStyle, textInsetStyle, verticalAlignStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
}

function renderPPTElementSVG(element: PPTElement) {
  if (element.kind === 'image') {
    return renderPPTImageSVG(element)
  }

  if (element.kind === 'line') {
    return renderPPTLineSVG(element)
  }

  if (element.kind === 'freeform') {
    return renderPPTFreeformSVG(element)
  }

  if (element.kind === 'table') {
    return renderPPTTableSVG(element)
  }

  if (element.kind === 'comment') {
    return renderPPTCommentSVG(element)
  }

  const attrs = getPPTElementSVGAttrs(element)
  const text = renderPPTTextBodySVG({
    body: element.textBody,
    geometry: element.geometry,
    inset: getPPTElementTextInset(element),
    style: element.style,
    verticalAlign: getPPTElementTextVerticalAlign(element),
  })

  if (element.kind === 'shape') {
    return `<g ${attrs}>${renderPPTShapeSVG(element)}${text}</g>`
  }

  return `<g ${attrs}>${text}</g>`
}

function renderPPTImageSVG(element: PPTImage) {
  const crop = element.crop ?? { x: 50, y: 50 }
  const fit = element.fit ?? 'cover'
  const cropRect = getPPTImageCropRect(element)
  const preserveAspectRatio = `${getPPTSvgImageAlignX(crop.x)}${getPPTSvgImageAlignY(crop.y)} ${fit === 'contain' ? 'meet' : 'slice'}`
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-image-fit="${fit}"`,
    `data-ppt-image-crop-x="${formatNumber(crop.x)}"`,
    `data-ppt-image-crop-y="${formatNumber(crop.y)}"`,
    getPPTImageCropRectSVGAttrs(element),
    getPPTImageAdjustmentsSVGAttrs(element),
    getPPTImageClipShapeSVGAttr(element),
  ].join(' ')
  const clipShape = getPPTImageClipShape(element)
  const shouldClip = Boolean(cropRect || clipShape)
  const stroke = renderPPTImageStrokeSVG(element)

  if (shouldClip) {
    const clipId = `ppt-svg-image-clip-${escapeHtml(element.id)}`
    const image = cropRect
      ? `<image href="${escapeHtml(element.src)}" x="${formatNumber(element.geometry.x - element.geometry.w * cropRect.left / cropRect.visibleWidth)}" y="${formatNumber(element.geometry.y - element.geometry.h * cropRect.top / cropRect.visibleHeight)}" width="${formatNumber(element.geometry.w * 100 / cropRect.visibleWidth)}" height="${formatNumber(element.geometry.h * 100 / cropRect.visibleHeight)}" preserveAspectRatio="none"><title>${escapeHtml(getPPTImageAltText(element))}</title></image>`
      : `<image href="${escapeHtml(element.src)}" x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" preserveAspectRatio="${preserveAspectRatio}"><title>${escapeHtml(getPPTImageAltText(element))}</title></image>`
    const defs = `<defs><clipPath id="${clipId}">${renderPPTImageClipPathSVG(element)}</clipPath></defs>`

    return `<g ${attrs}>${defs}<g clip-path="url(#${clipId})">${image}</g>${stroke}</g>`
  }

  return `<g ${attrs}><image href="${escapeHtml(element.src)}" x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" preserveAspectRatio="${preserveAspectRatio}"><title>${escapeHtml(getPPTImageAltText(element))}</title></image>${stroke}</g>`
}

function getPPTImageCropRectHTMLAttrs(element: PPTImage) {
  const cropRect = getPPTImageCropRectData(element)

  return cropRect
    ? [
        `data-ppt-image-crop-left="${formatNumber(cropRect.left)}"`,
        `data-ppt-image-crop-right="${formatNumber(cropRect.right)}"`,
        `data-ppt-image-crop-top="${formatNumber(cropRect.top)}"`,
        `data-ppt-image-crop-bottom="${formatNumber(cropRect.bottom)}"`,
      ].map((attr) => ` ${attr}`).join('')
    : ''
}

function getPPTImageAdjustmentsHTMLAttrs(element: PPTImage) {
  const attrs = getPPTImageAdjustmentAttrEntries(element)

  return attrs.length > 0
    ? ` ${attrs.join(' ')}`
    : ''
}

function getPPTImageAdjustmentsSVGAttrs(element: PPTImage) {
  return getPPTImageAdjustmentAttrEntries(element).join(' ')
}

function getPPTImageClipShapeHTMLAttr(element: PPTImage) {
  return element.clipShape
    ? ` data-ppt-image-clip-shape="${escapeHtml(element.clipShape)}"`
    : ''
}

function getPPTImageClipShapeSVGAttr(element: PPTImage) {
  return element.clipShape
    ? `data-ppt-image-clip-shape="${escapeHtml(element.clipShape)}"`
    : ''
}

function getPPTImageClipShapeStyle(element: PPTImage) {
  if (element.clipShape === 'ellipse') {
    return 'border-radius:999px'
  }

  if (element.clipShape === 'diamond') {
    return 'clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)'
  }

  return ''
}

function getPPTImageStrokeStyle(element: PPTImage) {
  const stroke = getPPTElementStroke(element)

  return stroke
    ? `border:${formatNumber(stroke.width)}px solid ${escapeHtml(stroke.color)};border-style:${getPPTStrokeDashBorderStyle(stroke)}`
    : ''
}

function getPPTImageClipShape(element: PPTImage) {
  return element.clipShape === 'ellipse' || element.clipShape === 'diamond'
    ? element.clipShape
    : undefined
}

function renderPPTImageClipPathSVG(element: PPTImage) {
  const shape = getPPTImageClipShape(element)

  if (shape === 'ellipse') {
    return `<ellipse cx="${formatNumber(element.geometry.x + element.geometry.w / 2)}" cy="${formatNumber(element.geometry.y + element.geometry.h / 2)}" rx="${formatNumber(element.geometry.w / 2)}" ry="${formatNumber(element.geometry.h / 2)}"></ellipse>`
  }

  if (shape === 'diamond') {
    const x = element.geometry.x
    const y = element.geometry.y
    const w = element.geometry.w
    const h = element.geometry.h

    return `<polygon points="${formatNumber(x + w / 2)},${formatNumber(y)} ${formatNumber(x + w)},${formatNumber(y + h / 2)} ${formatNumber(x + w / 2)},${formatNumber(y + h)} ${formatNumber(x)},${formatNumber(y + h / 2)}"></polygon>`
  }

  return `<rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}"></rect>`
}

function renderPPTImageStrokeSVG(element: PPTImage) {
  const stroke = getPPTElementStroke(element)

  if (!stroke) {
    return ''
  }

  const attrs = `fill="none" stroke="${escapeHtml(stroke.color)}" stroke-width="${formatNumber(stroke.width)}"${getPPTStrokeDashArraySvgAttr(stroke)}`
  const shape = getPPTImageClipShape(element)

  if (shape === 'ellipse') {
    return `<ellipse cx="${formatNumber(element.geometry.x + element.geometry.w / 2)}" cy="${formatNumber(element.geometry.y + element.geometry.h / 2)}" rx="${formatNumber(element.geometry.w / 2)}" ry="${formatNumber(element.geometry.h / 2)}" ${attrs}></ellipse>`
  }

  if (shape === 'diamond') {
    const x = element.geometry.x
    const y = element.geometry.y
    const w = element.geometry.w
    const h = element.geometry.h

    return `<polygon points="${formatNumber(x + w / 2)},${formatNumber(y)} ${formatNumber(x + w)},${formatNumber(y + h / 2)} ${formatNumber(x + w / 2)},${formatNumber(y + h)} ${formatNumber(x)},${formatNumber(y + h / 2)}" ${attrs}></polygon>`
  }

  return `<rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" ${attrs}></rect>`
}

function getPPTImageAdjustmentAttrEntries(element: PPTImage) {
  const adjustments = element.adjustments

  if (!adjustments) {
    return []
  }

  return [
    adjustments.grayscale === true
      ? 'data-ppt-image-adjustment-grayscale="true"'
      : '',
    adjustments.brightness === undefined
      ? ''
      : `data-ppt-image-adjustment-brightness="${formatNumber(adjustments.brightness)}"`,
    adjustments.contrast === undefined
      ? ''
      : `data-ppt-image-adjustment-contrast="${formatNumber(adjustments.contrast)}"`,
  ].filter(Boolean)
}

function getPPTImageCropRectSVGAttrs(element: PPTImage) {
  const cropRect = getPPTImageCropRectData(element)

  return cropRect
    ? [
        `data-ppt-image-crop-left="${formatNumber(cropRect.left)}"`,
        `data-ppt-image-crop-right="${formatNumber(cropRect.right)}"`,
        `data-ppt-image-crop-top="${formatNumber(cropRect.top)}"`,
        `data-ppt-image-crop-bottom="${formatNumber(cropRect.bottom)}"`,
      ].join(' ')
    : ''
}

function getPPTImageCropRectData(element: PPTImage) {
  const crop = element.crop

  if (!crop) {
    return null
  }

  const left = crop.left ?? 0
  const right = crop.right ?? 0
  const top = crop.top ?? 0
  const bottom = crop.bottom ?? 0

  return left !== 0 || right !== 0 || top !== 0 || bottom !== 0
    ? { bottom, left, right, top }
    : null
}

function getPPTImageCropRect(element: PPTImage) {
  const rect = getPPTImageCropRectData(element)

  if (!rect ||
    rect.left < 0 ||
    rect.right < 0 ||
    rect.top < 0 ||
    rect.bottom < 0
  ) {
    return null
  }

  const visibleWidth = 100 - rect.left - rect.right
  const visibleHeight = 100 - rect.top - rect.bottom

  return visibleWidth > 0 && visibleHeight > 0
    ? { ...rect, visibleHeight, visibleWidth }
    : null
}

function renderPPTShapeSVG(element: PPTShape) {
  const fill = element.fill.color
  const fillOpacity = getPPTFillOpacitySvgAttr(element.fill)
  const stroke = element.stroke
    ? ` stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"${getPPTStrokeDashArraySvgAttr(element.stroke)}`
    : ' stroke="none"'

  if (element.shape === 'ellipse') {
    return `<ellipse cx="${formatNumber(element.geometry.x + element.geometry.w / 2)}" cy="${formatNumber(element.geometry.y + element.geometry.h / 2)}" rx="${formatNumber(element.geometry.w / 2)}" ry="${formatNumber(element.geometry.h / 2)}" fill="${escapeHtml(fill)}"${fillOpacity}${stroke} />`
  }

  if (element.shape === 'diamond') {
    const left = element.geometry.x
    const top = element.geometry.y
    const right = element.geometry.x + element.geometry.w
    const bottom = element.geometry.y + element.geometry.h
    const centerX = element.geometry.x + element.geometry.w / 2
    const centerY = element.geometry.y + element.geometry.h / 2
    const points = [
      `${formatNumber(centerX)},${formatNumber(top)}`,
      `${formatNumber(right)},${formatNumber(centerY)}`,
      `${formatNumber(centerX)},${formatNumber(bottom)}`,
      `${formatNumber(left)},${formatNumber(centerY)}`,
    ].join(' ')

    return `<polygon points="${points}" fill="${escapeHtml(fill)}"${fillOpacity}${stroke} />`
  }

  const cornerRadius = formatNumber(getPPTShapeCornerRadius(element))

  return `<rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" rx="${cornerRadius}" ry="${cornerRadius}" fill="${escapeHtml(fill)}"${fillOpacity}${stroke} />`
}

function renderPPTLineSVG(element: PPTLine) {
  const markerId = `ppt-svg-line-marker-${escapeHtml(element.id)}`
  const marker = element.startMarker === 'arrow' || element.endMarker === 'arrow'
    ? `<defs><marker id="${markerId}" markerHeight="8" markerUnits="strokeWidth" markerWidth="8" orient="auto-start-reverse" refX="7" refY="4" viewBox="0 0 8 8"><path d="M 0 0 L 8 4 L 0 8 z" fill="${escapeHtml(element.stroke.color)}"></path></marker></defs>`
    : ''
  const markerStart = element.startMarker === 'arrow'
    ? ` marker-start="url(#${markerId})"`
    : ''
  const markerEnd = element.endMarker === 'arrow'
    ? ` marker-end="url(#${markerId})"`
    : ''
  const route = element.route ?? 'straight'
  const lineMarkup = route === 'elbow'
    ? `<path data-ppt-line-path d="${getPPTLineSVGPath(element)}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round" stroke-linejoin="round"${markerStart}${markerEnd}></path>`
    : renderPPTStraightLineSVG(element, markerStart, markerEnd)
  const connectionAttrs = [
    `data-ppt-line-route="${route}"`,
    element.startConnection
      ? `data-ppt-start-connection="${escapeHtml(element.startConnection.elementId)}:${element.startConnection.anchor}"`
      : '',
    element.endConnection
      ? `data-ppt-end-connection="${escapeHtml(element.endConnection.elementId)}:${element.endConnection.anchor}"`
      : '',
  ].filter(Boolean).join(' ')

  return `<g ${getPPTElementSVGAttrs(element)} ${connectionAttrs}>${marker}${lineMarkup}</g>`
}

function renderPPTStraightLineSVG(
  element: PPTLine,
  markerStart: string,
  markerEnd: string,
) {
  const start = getPPTLineWorldPoint(element, element.start)
  const end = getPPTLineWorldPoint(element, element.end)

  return `<line x1="${formatNumber(start.x)}" y1="${formatNumber(start.y)}" x2="${formatNumber(end.x)}" y2="${formatNumber(end.y)}" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round"${markerStart}${markerEnd}></line>`
}

function renderPPTTextBodySVG({
  body,
  geometry,
  inset,
  style,
  verticalAlign,
}: {
  body: PPTTextBody | undefined
  geometry: PPTElement['geometry']
  inset: PPTTextInset
  style: PPTTextStyle | undefined
  verticalAlign: PPTTextVerticalAlign
}) {
  if (!body) {
    return ''
  }

  const fontSize = style?.fontSize ?? 24
  let y = geometry.y + inset.top + getPPTTextVerticalAlignOffset({
    body,
    fontSize,
    geometry,
    inset,
    verticalAlign,
  })

  let numberedIndex = 0

  return body.paragraphs.map((paragraph) => {
    const bullet = paragraph.bullet === 'bullet'
    const numbered = paragraph.bullet === 'numbered'
    const listLevel = getPPTParagraphListLevel(paragraph)
    const runs = paragraph.runs.map(renderPPTTextRunSVG).join('')
    const listPrefix = bullet
      ? '<tspan data-ppt-bullet="true">&#8226; </tspan>'
      : numbered
        ? `<tspan data-ppt-numbered="true">${numberedIndex + 1}. </tspan>`
        : ''
    const lineHeight = getPPTParagraphLineHeight(paragraph)
    const spacingBefore = getPPTParagraphSpacingBefore(paragraph)
    const spacingAfter = getPPTParagraphSpacingAfter(paragraph)
    const textAnchor = getPPTSvgTextAnchor(paragraph.align)
    const x = getPPTSvgTextX({
      align: paragraph.align,
      geometry,
      inset,
    }) + fontSize * getSlideEditTextParagraphListLevelIndentEm(listLevel)
    y += spacingBefore + fontSize
    if (numbered) {
      numberedIndex += 1
    }
    const attrs = [
      'class="ppt-svg-text-paragraph"',
      bullet ? 'data-ppt-bullet="true"' : '',
      numbered ? 'data-ppt-numbered="true"' : '',
      paragraph.bullet ? `data-ppt-list="${escapeHtml(paragraph.bullet)}"` : '',
      `data-ppt-list-level="${formatNumber(listLevel)}"`,
      `data-ppt-line-height="${formatNumber(lineHeight)}"`,
      `data-ppt-spacing-after="${formatNumber(spacingAfter)}"`,
      `data-ppt-spacing-before="${formatNumber(spacingBefore)}"`,
      `data-ppt-text-inset="${escapeHtml(formatPPTTextInsetData(inset))}"`,
      `data-ppt-vertical-align="${escapeHtml(verticalAlign)}"`,
      `x="${formatNumber(x)}"`,
      `y="${formatNumber(y)}"`,
      `fill="${escapeHtml(style?.color ?? '#111827')}"`,
      `data-ppt-font-family="${escapeHtml(normalizePPTTextFontFamily(style?.fontFamily))}"`,
      `font-family="${escapeHtml(getPPTTextFontFamilyCSS(style?.fontFamily))}"`,
      `font-size="${formatNumber(fontSize)}"`,
      `font-weight="${getPPTSvgFontWeight(style)}"`,
      `text-anchor="${textAnchor}"`,
    ].filter(Boolean).join(' ')

    y += fontSize * lineHeight + spacingAfter

    return `<text ${attrs}>${listPrefix}${runs}</text>`
  }).join('')
}

function renderPPTTextRunSVG(run: PPTRun) {
  const textDecoration = getPPTTextRunTextDecoration(run)
  const attrs = [
    run.italic === true ? 'data-ppt-run-italic="true" font-style="italic"' : '',
    run.fontFamily
      ? `data-ppt-run-font-family="${escapeHtml(run.fontFamily)}" font-family="${escapeHtml(run.fontFamily)}"`
      : '',
    run.hyperlink ? `data-ppt-run-hyperlink-url="${escapeHtml(run.hyperlink.url)}"` : '',
    run.underline === true ? 'data-ppt-run-underline="true"' : '',
    run.strikethrough === true ? 'data-ppt-run-strikethrough="true"' : '',
    textDecoration ? `text-decoration="${textDecoration}"` : '',
    run.bold === true ? 'font-weight="700"' : '',
    run.color ? `fill="${escapeHtml(run.color)}"` : '',
    run.highlight ? `data-ppt-run-highlight="${escapeHtml(run.highlight)}"` : '',
    run.size ? `font-size="${formatNumber(run.size)}"` : '',
  ].filter(Boolean).join(' ')

  return attrs
    ? `<tspan ${attrs}>${escapeHtml(run.text)}</tspan>`
    : `<tspan>${escapeHtml(run.text)}</tspan>`
}

function renderPPTSlideNotesHTML(slideId: string, notes: string | undefined) {
  const trimmed = notes?.trim()

  if (!trimmed) {
    return ''
  }

  return `  <aside class="ppt-notes" data-ppt-notes-for="${escapeHtml(slideId)}">${escapeHtml(notes ?? '')}</aside>`
}

function exportCSS() {
  return [
    ':root{font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#111827;background:#e5e7eb;}',
    '*{box-sizing:border-box;}',
    'body{margin:0;}',
    '@page{size:16in 9in;margin:0;}',
    `.ppt-deck{display:grid;gap:32px;padding:32px;}`,
    `.ppt-slide{position:relative;width:${PPT_SLIDE_WIDTH}px;height:${PPT_SLIDE_HEIGHT}px;overflow:hidden;background:#fff;break-after:page;}`,
    '.ppt-element{position:absolute;margin:0;overflow:hidden;white-space:pre-wrap;overflow-wrap:anywhere;display:flex;align-items:center;padding:18px;}',
    '.ppt-image{display:block;object-fit:cover;padding:0;}',
    '.ppt-freeform{display:block;overflow:visible;padding:0;}',
    '.ppt-freeform-text{position:absolute;inset:0;display:flex;overflow:hidden;}',
    '.ppt-line{display:block;overflow:visible;padding:0;}',
    '.ppt-comment{display:grid;grid-template-rows:auto minmax(0,1fr);padding:0;border:1px solid #d97706;border-radius:8px;background:#fffbeb;color:#78350f;box-shadow:0 10px 22px rgb(120 53 15 / 18%);}',
    '.ppt-comment[data-ppt-comment-resolved="true"]{opacity:.62;}',
    '.ppt-comment-meta{display:flex;gap:6px;align-items:center;min-width:0;padding:8px 10px 6px;border-bottom:1px solid #fde68a;color:#92400e;font-size:12px;font-weight:700;line-height:1;}',
    '.ppt-comment-created{margin-left:auto;color:#b45309;font-weight:600;}',
    '.ppt-comment-body{margin:0;padding:9px 10px 12px;overflow:hidden;font-size:17px;font-weight:650;line-height:1.22;white-space:pre-wrap;}',
    '.ppt-table{display:table;table-layout:fixed;border-collapse:collapse;padding:0;background:#fff;color:#111827;font-size:18px;line-height:1.15;}',
    '.ppt-table th,.ppt-table td{height:1px;padding:8px 10px;overflow:hidden;border:1px solid #dbe3ef;text-align:left;text-overflow:ellipsis;white-space:nowrap;}',
    '.ppt-table th{background:#eff6ff;font-weight:700;}',
    '.ppt-text{align-items:flex-start;padding:0;}',
    '.ppt-text-paragraph{--ppt-paragraph-list-level-indent:0em;display:block;min-height:1em;padding-left:var(--ppt-paragraph-list-level-indent);}',
    '.ppt-text-paragraph[data-ppt-bullet="true"]{position:relative;padding-left:calc(var(--ppt-paragraph-list-level-indent) + 1.1em);}',
    '.ppt-text-paragraph[data-ppt-bullet="true"]::before{content:"\\2022";position:absolute;left:var(--ppt-paragraph-list-level-indent);}',
    '.ppt-element{counter-reset:ppt-numbered-list;}',
    '.ppt-text-paragraph[data-ppt-numbered="true"]{position:relative;padding-left:calc(var(--ppt-paragraph-list-level-indent) + 1.45em);counter-increment:ppt-numbered-list;}',
    '.ppt-text-paragraph[data-ppt-numbered="true"]::before{content:counter(ppt-numbered-list) ".";position:absolute;left:var(--ppt-paragraph-list-level-indent);}',
    '.ppt-shape{border-radius:24px;}',
    '.ppt-shape-ellipse{border-radius:999px;}',
    '.ppt-shape-diamond{clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%);}',
    `.ppt-notes{width:${PPT_SLIDE_WIDTH}px;padding:16px 20px;background:#fff;color:#344054;font-size:16px;line-height:1.4;white-space:pre-wrap;}`,
    '@media print{body{background:#fff;}.ppt-deck{display:block;padding:0;}.ppt-slide{break-after:page;page-break-after:always;}}',
  ].join('\n')
}

function renderPPTTextBodyHTML(body: PPTTextBody | undefined) {
  if (!body) {
    return ''
  }

  return body.paragraphs.map((paragraph) => {
    const runs = paragraph.runs
      .map(renderPPTTextRunHTML)
      .join('')
    const bulletAttr = getPPTParagraphListHTMLAttrs(paragraph)
    const paragraphAttrs = getPPTParagraphHTMLAttrs(paragraph)
    const paragraphStyleAttr = getPPTParagraphStyleAttr(paragraph)

    return `<span class="ppt-text-paragraph"${bulletAttr}${paragraphAttrs}${paragraphStyleAttr}>${runs}</span>`
  }).join('')
}

function renderPPTTextRunHTML(run: PPTRun) {
  const attrs = [
    run.italic === true ? 'data-ppt-run-italic="true"' : '',
    run.fontFamily
      ? `data-ppt-run-font-family="${escapeHtml(run.fontFamily)}"`
      : '',
    run.highlight ? `data-ppt-run-highlight="${escapeHtml(run.highlight)}"` : '',
    run.hyperlink ? `data-ppt-run-hyperlink-url="${escapeHtml(run.hyperlink.url)}"` : '',
    run.strikethrough === true ? 'data-ppt-run-strikethrough="true"' : '',
    run.underline === true ? 'data-ppt-run-underline="true"' : '',
    renderPPTTextRunStyleAttr(run),
  ].filter(Boolean).join(' ')

  return attrs
    ? `<span ${attrs}>${escapeHtml(run.text)}</span>`
    : escapeHtml(run.text)
}

function renderPPTTextRunStyleAttr(run: PPTRun) {
  const styles = [
    run.bold === true ? 'font-weight:700' : '',
    run.highlight ? `background-color:${escapeHtml(run.highlight)}` : '',
    run.color ? `color:${escapeHtml(run.color)}` : '',
    run.fontFamily
      ? `font-family:${escapeHtml(formatPPTTextRunFontFamilyCSS(run.fontFamily))}`
      : '',
    run.italic === true ? 'font-style:italic' : '',
    run.size ? `font-size:${run.size}px` : '',
    getPPTTextRunTextDecoration(run)
      ? `text-decoration:${getPPTTextRunTextDecoration(run)}`
      : '',
  ].filter(Boolean).join(';')

  return styles ? `style="${styles}"` : ''
}

function formatPPTTextRunFontFamilyCSS(fontFamily: string) {
  return `"${fontFamily
    .trim()
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/[\n\r\f]/g, ' ')}"`
}

function getPPTTextRunTextDecoration(
  run: Pick<PPTRun, 'strikethrough' | 'underline'>,
) {
  const decorations = [
    run.underline === true ? 'underline' : '',
    run.strikethrough === true ? 'line-through' : '',
  ].filter(Boolean)

  return decorations.length > 0 ? decorations.join(' ') : undefined
}

function getPPTParagraphListHTMLAttrs(paragraph: PPTParagraph) {
  if (paragraph.bullet === 'bullet') {
    return ' data-ppt-bullet="true" data-ppt-list="bullet"'
  }

  if (paragraph.bullet === 'numbered') {
    return ' data-ppt-numbered="true" data-ppt-list="numbered"'
  }

  return ''
}

function hasPPTTextBodyBullet(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'bullet')
}

function hasPPTTextBodyNumbered(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'numbered')
}

function renderPPTLineHTML(element: PPTLine, style: string[]) {
  const markerId = `ppt-line-marker-${escapeHtml(element.id)}`
  const marker = element.startMarker === 'arrow' || element.endMarker === 'arrow'
    ? `<defs><marker id="${markerId}" markerHeight="8" markerUnits="strokeWidth" markerWidth="8" orient="auto-start-reverse" refX="7" refY="4" viewBox="0 0 8 8"><path d="M 0 0 L 8 4 L 0 8 z" fill="${escapeHtml(element.stroke.color)}"></path></marker></defs>`
    : ''
  const markerStart = element.startMarker === 'arrow'
    ? ` marker-start="url(#${markerId})"`
    : ''
  const markerEnd = element.endMarker === 'arrow'
    ? ` marker-end="url(#${markerId})"`
    : ''
  const route = element.route ?? 'straight'
  const lineMarkup = route === 'elbow'
    ? `<path data-ppt-line-path d="${getPPTLinePath(element)}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${element.stroke.width}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round" stroke-linejoin="round"${markerStart}${markerEnd}></path>`
    : `<line x1="${element.start.x}" y1="${element.start.y}" x2="${element.end.x}" y2="${element.end.y}" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${element.stroke.width}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round"${markerStart}${markerEnd}></line>`

  const connectionAttrs = [
    `data-ppt-line-route="${route}"`,
    element.startConnection
      ? `data-ppt-start-connection="${escapeHtml(element.startConnection.elementId)}:${element.startConnection.anchor}"`
      : '',
    element.endConnection
      ? `data-ppt-end-connection="${escapeHtml(element.endConnection.elementId)}:${element.endConnection.anchor}"`
      : '',
  ].filter(Boolean).join(' ')

  return `    <svg class="ppt-element ppt-line" data-ppt-element="${escapeHtml(element.id)}"${getPPTElementAltTextHTMLAttr(element)}${getPPTElementHyperlinkHTMLAttr(element)}${getPPTElementStrokeDashHTMLAttr(element)}${getPPTElementOpacityHTMLAttr(element)}${getPPTElementShadowHTMLAttrs(element)} ${connectionAttrs} style="${style.filter(Boolean).join(';')}" viewBox="0 0 ${element.geometry.w} ${element.geometry.h}" preserveAspectRatio="none" aria-hidden="true">${marker}${lineMarkup}</svg>`
}

function renderPPTFreeformHTML(element: PPTFreeform, style: string[]) {
  const fill = element.fill ? element.fill.color : 'none'
  const fillOpacity = element.fill ? getPPTFillOpacitySvgAttr(element.fill) : ''
  const attrs = [
    `data-ppt-element="${escapeHtml(element.id)}"`,
    `data-ppt-kind="freeform"`,
    `data-ppt-freeform-point-mode="${element.pointMode ?? 'freehand'}"`,
    `data-ppt-freeform-points="${element.points.length}"`,
    element.fill ? getPPTElementFillOpacityDataAttr(element.fill) : '',
    getPPTElementAltTextHTMLAttr(element).trim(),
    getPPTElementHyperlinkHTMLAttr(element).trim(),
    getPPTElementStrokeDashHTMLAttr(element).trim(),
    getPPTElementOpacityHTMLAttr(element).trim(),
    getPPTElementShadowHTMLAttrs(element).trim(),
  ].join(' ')
  const path = `<path data-ppt-freeform-path d="${escapeHtml(getPPTFreeformPathData(element))}" fill="${escapeHtml(fill)}"${fillOpacity} stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round" stroke-linejoin="round"></path>`

  if (!element.textBody) {
    return `    <svg class="ppt-element ppt-freeform" ${attrs} style="${style.filter(Boolean).join(';')}" viewBox="0 0 ${formatNumber(element.geometry.w)} ${formatNumber(element.geometry.h)}" preserveAspectRatio="none" aria-hidden="true">${path}</svg>`
  }

  const text = renderPPTTextBodyHTML(element.textBody)
  const textStyle = element.style ? exportTextStyle(element.style) : ''
  const paragraphStyle = `text-align:${element.textBody.paragraphs[0]?.align ?? 'left'}`
  const textInset = getPPTElementTextInset(element)
  const textInsetStyle = `padding:${getPPTTextInsetCSS(textInset)}`
  const verticalAlign = getPPTElementTextVerticalAlign(element)
  const verticalAlignStyle = `align-items:${getPPTTextVerticalAlignCSS(verticalAlign)}`

  return `    <div class="ppt-element ppt-freeform" ${attrs} style="${style.filter(Boolean).join(';')}"><svg style="position:absolute;inset:0;width:100%;height:100%;overflow:visible" viewBox="0 0 ${formatNumber(element.geometry.w)} ${formatNumber(element.geometry.h)}" preserveAspectRatio="none" aria-hidden="true">${path}</svg><div class="ppt-freeform-text" style="${[textStyle, textInsetStyle, verticalAlignStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div></div>`
}

function renderPPTFreeformSVG(element: PPTFreeform) {
  const fill = element.fill ? element.fill.color : 'none'
  const fillOpacity = element.fill ? getPPTFillOpacitySvgAttr(element.fill) : ''
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-freeform-points="${element.points.length}"`,
  ].join(' ')
  const text = renderPPTTextBodySVG({
    body: element.textBody,
    geometry: element.geometry,
    inset: getPPTElementTextInset(element),
    style: element.style,
    verticalAlign: getPPTElementTextVerticalAlign(element),
  })

  return `<g ${attrs}><path data-ppt-freeform-path d="${escapeHtml(getPPTFreeformWorldPathData(element))}" fill="${escapeHtml(fill)}"${fillOpacity} stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"${getPPTStrokeDashArraySvgAttr(element.stroke)} stroke-linecap="round" stroke-linejoin="round"></path>${text}</g>`
}

function renderPPTCommentHTML(
  element: PPTComment,
  style: string[],
  transformAttrs: string,
) {
  const author = element.authorName ?? 'You'
  const createdAt = element.createdAt ?? 'Just now'
  const thread = getPPTExportCommentThread(element)
  const resolvedAttr = element.resolved === true
    ? ' data-ppt-comment-resolved="true"'
    : ''
  const threadMarkup = thread.map((message, index) =>
    `<span data-ppt-comment-thread-message="${escapeHtml(message.id)}" data-ppt-comment-thread-message-index="${index}"><span data-ppt-comment-thread-author>${escapeHtml(message.authorName)}</span><span data-ppt-comment-thread-created>${escapeHtml(message.createdAt)}</span><span data-ppt-comment-thread-body>${escapeHtml(message.body)}</span></span>`
  ).join('')

  return `    <div class="ppt-element ppt-comment" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${getPPTElementAltTextHTMLAttr(element)}${getPPTElementHyperlinkHTMLAttr(element)}${getPPTElementOpacityHTMLAttr(element)}${getPPTElementShadowHTMLAttrs(element)}${resolvedAttr} data-ppt-comment-thread-count="${thread.length}" style="${style.filter(Boolean).join(';')}"><div class="ppt-comment-meta"><span>${escapeHtml(author)}</span><span class="ppt-comment-created">${escapeHtml(createdAt)}</span></div><p class="ppt-comment-body">${escapeHtml(element.body)}</p><div class="ppt-comment-thread" data-ppt-comment-thread-count="${thread.length}" hidden>${threadMarkup}</div></div>`
}

function renderPPTCommentSVG(element: PPTComment) {
  const attrs = [
    getPPTElementSVGAttrs(element),
    element.resolved === true ? 'data-ppt-comment-resolved="true"' : '',
    `data-ppt-comment-thread-count="${getPPTExportCommentThread(element).length}"`,
  ].filter(Boolean).join(' ')
  const x = element.geometry.x
  const y = element.geometry.y
  const metaHeight = Math.min(30, element.geometry.h * 0.28)
  const body = element.body || 'Comment'
  const author = element.authorName ?? 'You'

  return [
    `<g ${attrs}>`,
    `<rect x="${formatNumber(x)}" y="${formatNumber(y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" rx="8" fill="#fffbeb" stroke="#d97706" stroke-width="1"></rect>`,
    `<rect data-ppt-comment-meta="true" x="${formatNumber(x)}" y="${formatNumber(y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(metaHeight)}" rx="8" fill="#fef3c7"></rect>`,
    `<text data-ppt-comment-author="true" x="${formatNumber(x + 10)}" y="${formatNumber(y + metaHeight / 2)}" fill="#92400e" font-family="Inter, Arial, sans-serif" font-size="12" font-weight="700" dominant-baseline="middle">${escapeHtml(author)}</text>`,
    `<text data-ppt-comment-body="true" x="${formatNumber(x + 10)}" y="${formatNumber(y + metaHeight + 24)}" fill="#78350f" font-family="Inter, Arial, sans-serif" font-size="17" font-weight="650">${escapeHtml(body)}</text>`,
    '</g>',
  ].join('')
}

function getPPTExportCommentThread(comment: PPTComment): PPTCommentThreadMessage[] {
  if (comment.thread && comment.thread.length > 0) {
    return comment.thread
  }

  return [{
    authorName: comment.authorName ?? 'You',
    body: comment.body,
    createdAt: comment.createdAt ?? 'Just now',
    id: `${comment.id}:message-1`,
  }]
}

function renderPPTTableHTML(
  element: PPTTable,
  style: string[],
  transformAttrs: string,
) {
  const columnCount = getPPTTableColumnCount(element.rows)
  const rowCount = element.rows.length
  const columnWidths = getPPTTableResolvedColumnWidths(element)
  const rowHeights = getPPTTableResolvedRowHeights(element)
  const colGroup = columnWidths.length > 0
    ? `<colgroup>${columnWidths.map((width) =>
        `<col style="width:${formatNumber(width)}px">`).join('')}</colgroup>`
    : ''
  const head = element.rows[0]
    ? `<thead>${renderPPTTableHTMLRow({
        columnCount,
        element,
        row: element.rows[0],
        rowHeight: rowHeights[0],
        rowIndex: 0,
        tagName: 'th',
      })}</thead>`
    : ''
  const bodyRows = element.rows
    .slice(1)
    .map((row, index) =>
      renderPPTTableHTMLRow({
        columnCount,
        element,
        row,
        rowHeight: rowHeights[index + 1],
        rowIndex: index + 1,
        tagName: 'td',
      }))
    .join('')
  const body = `<tbody>${bodyRows}</tbody>`
  const attrs = [
    `class="ppt-element ppt-table"`,
    `data-ppt-element="${escapeHtml(element.id)}"`,
    transformAttrs.trim(),
    getPPTElementAltTextHTMLAttr(element).trim(),
    getPPTElementHyperlinkHTMLAttr(element).trim(),
    getPPTElementOpacityHTMLAttr(element).trim(),
    getPPTElementShadowHTMLAttrs(element).trim(),
    `data-ppt-table-rows="${rowCount}"`,
    `data-ppt-table-cols="${columnCount}"`,
    `data-ppt-table-column-widths="${escapeHtml(formatPPTTableTrackSizesAttribute(columnWidths))}"`,
    `data-ppt-table-row-heights="${escapeHtml(formatPPTTableTrackSizesAttribute(rowHeights))}"`,
    `style="${style.filter(Boolean).join(';')}"`,
  ].filter(Boolean).join(' ')

  return `    <table ${attrs}>${colGroup}${head}${body}</table>`
}

function renderPPTTableHTMLRow({
  columnCount,
  element,
  row,
  rowHeight,
  rowIndex,
  tagName,
}: {
  columnCount: number
  element: PPTTable
  row: readonly string[]
  rowHeight: number | undefined
  rowIndex: number
  tagName: 'td' | 'th'
}) {
  const rowStyle = rowHeight === undefined
    ? ''
    : ` style="height:${formatNumber(rowHeight)}px"`
  const cells = Array.from({ length: columnCount }, (_, index) => {
    if (isPPTTableCellHidden(element, rowIndex, index)) {
      return ''
    }

    const borders = getPPTTableCellBorders(element, rowIndex, index)
    const colSpan = getPPTTableCellColSpan(element, rowIndex, index)
    const fill = getPPTTableCellFill(element, rowIndex, index)
    const rowSpan = getPPTTableCellRowSpan(element, rowIndex, index)
    const textStyle = getPPTTableCellTextStyle(element, rowIndex, index)
    const cellAttrs = fill || textStyle || borders || colSpan > 1 || rowSpan > 1
      ? [
          borders?.bottom
            ? ` data-ppt-table-cell-border-bottom="${escapeHtml(formatPPTTableCellBorderData(borders.bottom))}"`
            : '',
          borders?.left
            ? ` data-ppt-table-cell-border-left="${escapeHtml(formatPPTTableCellBorderData(borders.left))}"`
            : '',
          borders?.right
            ? ` data-ppt-table-cell-border-right="${escapeHtml(formatPPTTableCellBorderData(borders.right))}"`
            : '',
          borders?.top
            ? ` data-ppt-table-cell-border-top="${escapeHtml(formatPPTTableCellBorderData(borders.top))}"`
            : '',
          colSpan > 1
            ? ` data-ppt-table-cell-col-span="${colSpan}" colspan="${colSpan}"`
            : '',
          fill
            ? ` data-ppt-table-cell-fill="${escapeHtml(fill.color)}"`
            : '',
          fill?.opacity === undefined
            ? ''
            : ` data-ppt-table-cell-fill-opacity="${escapeHtml(formatPPTFillOpacity(getPPTFillOpacity(fill)))}"`,
          rowSpan > 1
            ? ` data-ppt-table-cell-row-span="${rowSpan}" rowspan="${rowSpan}"`
            : '',
          textStyle?.align
            ? ` data-ppt-table-cell-align="${escapeHtml(textStyle.align)}"`
            : '',
          textStyle?.color
            ? ` data-ppt-table-cell-text-color="${escapeHtml(textStyle.color)}"`
            : '',
          textStyle?.fontSize
            ? ` data-ppt-table-cell-font-size="${escapeHtml(formatNumber(textStyle.fontSize))}"`
            : '',
          textStyle?.fontWeight
            ? ` data-ppt-table-cell-font-weight="${escapeHtml(textStyle.fontWeight)}"`
            : '',
          textStyle?.textInset
            ? ` data-ppt-table-cell-text-inset="${escapeHtml(formatPPTTextInsetData(textStyle.textInset))}"`
            : '',
          textStyle?.verticalAlign
            ? ` data-ppt-table-cell-vertical-align="${escapeHtml(textStyle.verticalAlign)}"`
            : '',
          formatPPTTableCellHTMLStyle(fill, textStyle, borders),
        ].join('')
      : ''

    return `<${tagName} data-ppt-table-cell="${index}"${cellAttrs}>${escapeHtml(row[index] ?? '')}</${tagName}>`
  }).join('')

  return `<tr${rowStyle}>${cells}</tr>`
}

function renderPPTTableSVG(element: PPTTable) {
  const columnCount = getPPTTableColumnCount(element.rows)
  const rowCount = element.rows.length
  const columnWidths = getPPTTableResolvedColumnWidths(element)
  const rowHeights = getPPTTableResolvedRowHeights(element)
  let rowY = element.geometry.y
  const cells = element.rows.flatMap((row, rowIndex) => {
    const cellHeight = rowHeights[rowIndex] ?? (rowCount > 0 ? element.geometry.h / rowCount : element.geometry.h)
    const y = rowY
    let columnX = element.geometry.x
    rowY += cellHeight

    return Array.from({ length: columnCount }, (_, columnIndex) => {
      const cellWidth = columnWidths[columnIndex] ?? (columnCount > 0 ? element.geometry.w / columnCount : element.geometry.w)
      const x = columnX
      const colSpan = getPPTTableCellColSpan(element, rowIndex, columnIndex)
      const rowSpan = getPPTTableCellRowSpan(element, rowIndex, columnIndex)
      const spannedCellHeight = sumPPTTableTrackSizes(rowHeights, rowIndex, rowSpan, cellHeight)
      const spannedCellWidth = sumPPTTableTrackSizes(columnWidths, columnIndex, colSpan, cellWidth)
      const fontSize = Math.max(10, Math.min(18, spannedCellHeight * 0.38))
      const headerAttrs = rowIndex === 0
        ? ' data-ppt-table-header="true"'
        : ''

      columnX += cellWidth

      if (isPPTTableCellHidden(element, rowIndex, columnIndex)) {
        return ''
      }

      const cellFill = getPPTTableCellFill(element, rowIndex, columnIndex)
      const cellBorders = getPPTTableCellBorders(element, rowIndex, columnIndex)
      const cellTextStyle = getPPTTableCellTextStyle(element, rowIndex, columnIndex)
      const fillAttrs = cellFill
        ? `fill="${escapeHtml(cellFill.color)}"${getPPTFillOpacitySvgAttr(cellFill)}`
        : `fill="${rowIndex === 0 ? '#eff6ff' : '#ffffff'}"`
      const textColor = cellTextStyle?.color ?? '#111827'
      const textFontSize = cellTextStyle?.fontSize ?? fontSize
      const textInset = getPPTTableCellTextInset({
        cellHeight: spannedCellHeight,
        cellWidth: spannedCellWidth,
        textInset: cellTextStyle?.textInset,
      })
      const textAnchor = getPPTTableCellTextAnchor(cellTextStyle?.align)
      const textBaseline = getPPTTableCellDominantBaseline(cellTextStyle?.verticalAlign)
      const textInsetAttr = cellTextStyle?.textInset
        ? ` data-ppt-table-cell-text-inset="${escapeHtml(formatPPTTextInsetData(textInset))}"`
        : ''
      const textVerticalAlignAttr = cellTextStyle?.verticalAlign
        ? ` data-ppt-table-cell-vertical-align="${escapeHtml(cellTextStyle.verticalAlign)}"`
        : ''
      const textX = getPPTTableCellTextX({
        align: cellTextStyle?.align,
        cellWidth: spannedCellWidth,
        inset: textInset,
        x,
      })
      const textY = getPPTTableCellTextY({
        cellHeight: spannedCellHeight,
        inset: textInset,
        verticalAlign: cellTextStyle?.verticalAlign,
        y,
      })
      const textFontWeight = cellTextStyle?.fontWeight
        ? getPPTTableCellFontWeightSvgAttr(cellTextStyle.fontWeight)
        : rowIndex === 0 ? ' font-weight="700"' : ''
      const borderLines = renderPPTTableCellBorderSVG({
        borders: cellBorders,
        cellHeight: spannedCellHeight,
        cellWidth: spannedCellWidth,
        x,
        y,
      })

      return [
        `<rect data-ppt-table-cell="${rowIndex}:${columnIndex}"${colSpan > 1 ? ` data-ppt-table-cell-col-span="${colSpan}"` : ''}${rowSpan > 1 ? ` data-ppt-table-cell-row-span="${rowSpan}"` : ''} x="${formatNumber(x)}" y="${formatNumber(y)}" width="${formatNumber(spannedCellWidth)}" height="${formatNumber(spannedCellHeight)}" ${fillAttrs} stroke="#dbe3ef" stroke-width="1"></rect>`,
        borderLines,
        `<text data-ppt-table-text="${rowIndex}:${columnIndex}"${headerAttrs}${textInsetAttr}${textVerticalAlignAttr} x="${formatNumber(textX)}" y="${formatNumber(textY)}" fill="${escapeHtml(textColor)}" font-family="Inter, Arial, sans-serif" font-size="${formatNumber(textFontSize)}"${textFontWeight} text-anchor="${textAnchor}" dominant-baseline="${textBaseline}">${escapeHtml(row[columnIndex] ?? '')}</text>`,
      ].join('')
    })
  }).join('')
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-table-rows="${rowCount}"`,
    `data-ppt-table-cols="${columnCount}"`,
    `data-ppt-table-column-widths="${escapeHtml(formatPPTTableTrackSizesAttribute(columnWidths))}"`,
    `data-ppt-table-row-heights="${escapeHtml(formatPPTTableTrackSizesAttribute(rowHeights))}"`,
  ].join(' ')

  return `<g ${attrs}><rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" fill="#ffffff" stroke="#94a3b8" stroke-width="1"></rect>${cells}</g>`
}

function formatPPTTableTrackSizesAttribute(trackSizes: readonly number[]) {
  return trackSizes.map((size) => Math.round(size)).join(' ')
}

function sumPPTTableTrackSizes(
  trackSizes: readonly number[],
  startIndex: number,
  span: number,
  fallbackSize: number,
) {
  if (span <= 1) {
    return fallbackSize
  }

  return trackSizes
    .slice(startIndex, startIndex + span)
    .reduce((total, size) => total + size, 0) || fallbackSize
}

function formatPPTTableCellHTMLStyle(
  fill: PPTFill | undefined,
  textStyle: PPTTableCellTextStyle | undefined,
  borders: PPTTableCellBorders | undefined,
) {
  const styles = [
    borders?.bottom ? `border-bottom:${getPPTTableCellBorderCSS(borders.bottom)}` : '',
    borders?.left ? `border-left:${getPPTTableCellBorderCSS(borders.left)}` : '',
    borders?.right ? `border-right:${getPPTTableCellBorderCSS(borders.right)}` : '',
    borders?.top ? `border-top:${getPPTTableCellBorderCSS(borders.top)}` : '',
    fill ? `background:${getPPTFillColorCSS(fill)}` : '',
    textStyle?.align ? `text-align:${textStyle.align}` : '',
    textStyle?.color ? `color:${textStyle.color}` : '',
    textStyle?.fontSize ? `font-size:${formatNumber(textStyle.fontSize)}px` : '',
    textStyle?.fontWeight
      ? `font-weight:${getPPTTableCellFontWeightCSS(textStyle.fontWeight)}`
      : '',
    textStyle?.textInset ? `padding:${getPPTTextInsetCSS(textStyle.textInset)}` : '',
    textStyle?.verticalAlign ? `vertical-align:${textStyle.verticalAlign}` : '',
  ].filter(Boolean)

  return styles.length > 0
    ? ` style="${escapeHtml(styles.join(';'))}"`
    : ''
}

function getPPTTableCellBorderCSS(stroke: PPTStroke) {
  return `${formatNumber(stroke.width)}px ${getPPTStrokeDashBorderStyle(stroke)} ${stroke.color}`
}

function formatPPTTableCellBorderData(stroke: PPTStroke) {
  return `${stroke.color} ${formatNumber(stroke.width)} ${getPPTStrokeDash(stroke)}`
}

function renderPPTTableCellBorderSVG({
  borders,
  cellHeight,
  cellWidth,
  x,
  y,
}: {
  borders: PPTTableCellBorders | undefined
  cellHeight: number
  cellWidth: number
  x: number
  y: number
}) {
  if (!borders) {
    return ''
  }

  return [
    borders.top
      ? renderPPTTableCellBorderLineSVG({
          side: 'top',
          stroke: borders.top,
          x1: x,
          x2: x + cellWidth,
          y1: y,
          y2: y,
        })
      : '',
    borders.right
      ? renderPPTTableCellBorderLineSVG({
          side: 'right',
          stroke: borders.right,
          x1: x + cellWidth,
          x2: x + cellWidth,
          y1: y,
          y2: y + cellHeight,
        })
      : '',
    borders.bottom
      ? renderPPTTableCellBorderLineSVG({
          side: 'bottom',
          stroke: borders.bottom,
          x1: x,
          x2: x + cellWidth,
          y1: y + cellHeight,
          y2: y + cellHeight,
        })
      : '',
    borders.left
      ? renderPPTTableCellBorderLineSVG({
          side: 'left',
          stroke: borders.left,
          x1: x,
          x2: x,
          y1: y,
          y2: y + cellHeight,
        })
      : '',
  ].join('')
}

function renderPPTTableCellBorderLineSVG({
  side,
  stroke,
  x1,
  x2,
  y1,
  y2,
}: {
  side: keyof PPTTableCellBorders
  stroke: PPTStroke
  x1: number
  x2: number
  y1: number
  y2: number
}) {
  return `<line data-ppt-table-cell-border="${side}" x1="${formatNumber(x1)}" x2="${formatNumber(x2)}" y1="${formatNumber(y1)}" y2="${formatNumber(y2)}" stroke="${escapeHtml(stroke.color)}" stroke-width="${formatNumber(stroke.width)}"${getPPTStrokeDashArraySvgAttr(stroke)} stroke-linecap="square"></line>`
}

function getPPTTableCellFontWeightCSS(
  fontWeight: PPTTableCellTextStyle['fontWeight'],
) {
  if (fontWeight === 'bold') {
    return '700'
  }

  return fontWeight === 'semibold' ? '600' : '400'
}

function getPPTTableCellFontWeightSvgAttr(
  fontWeight: PPTTableCellTextStyle['fontWeight'],
) {
  return ` font-weight="${getPPTTableCellFontWeightCSS(fontWeight)}"`
}

function getPPTTableCellTextAnchor(align: PPTTableCellTextStyle['align']) {
  if (align === 'center') {
    return 'middle'
  }

  return align === 'right' ? 'end' : 'start'
}

function getPPTTableCellTextX({
  align,
  cellWidth,
  inset,
  x,
}: {
  align: PPTTableCellTextStyle['align']
  cellWidth: number
  inset: PPTTextInset
  x: number
}) {
  if (align === 'center') {
    return x + inset.left + Math.max(0, cellWidth - inset.left - inset.right) / 2
  }

  if (align === 'right') {
    return x + cellWidth - inset.right
  }

  return x + inset.left
}

function getPPTTableCellTextY({
  cellHeight,
  inset,
  verticalAlign,
  y,
}: {
  cellHeight: number
  inset: PPTTextInset
  verticalAlign: PPTTableCellTextStyle['verticalAlign']
  y: number
}) {
  if (verticalAlign === 'top') {
    return y + inset.top
  }

  if (verticalAlign === 'bottom') {
    return y + cellHeight - inset.bottom
  }

  return y + inset.top + Math.max(0, cellHeight - inset.top - inset.bottom) / 2
}

function getPPTTableCellDominantBaseline(
  verticalAlign: PPTTableCellTextStyle['verticalAlign'],
) {
  if (verticalAlign === 'top') {
    return 'hanging'
  }

  return verticalAlign === 'bottom' ? 'text-after-edge' : 'middle'
}

function getPPTTableCellTextInset({
  cellHeight,
  cellWidth,
  textInset,
}: {
  cellHeight: number
  cellWidth: number
  textInset: PPTTableCellTextStyle['textInset']
}): PPTTextInset {
  const horizontal = Math.min(10, cellWidth * 0.12)
  const vertical = Math.min(8, cellHeight * 0.2)

  return {
    bottom: textInset?.bottom ?? vertical,
    left: textInset?.left ?? horizontal,
    right: textInset?.right ?? horizontal,
    top: textInset?.top ?? vertical,
  }
}

function getPPTElementTransform(element: PPTElement) {
  return createPPTCanvasCssBoundsTransform({
    flipX: element.flipH === true,
    flipY: element.flipV === true,
    rotation: element.geometry.rotation,
  })
}

function getPPTElementTransformAttrs(element: PPTElement) {
  return [
    element.flipH === true ? ' data-ppt-flip-h="true"' : '',
    element.flipV === true ? ' data-ppt-flip-v="true"' : '',
    getPPTElementAnimationHTMLAttrs(element),
  ].join('')
}

function getPPTElementOpacityHTMLAttr(element: PPTElement) {
  return ` data-ppt-opacity="${escapeHtml(formatPPTElementOpacity(getPPTElementOpacity(element)))}"`
}

function getPPTElementAltTextHTMLAttr(element: PPTElement) {
  const altText = getPPTElementAltText(element)

  return altText
    ? ` data-ppt-alt-text="${escapeHtml(altText)}"`
    : ''
}

function getPPTElementAltTextSvgAttr(element: PPTElement) {
  const altText = getPPTElementAltText(element)

  return altText
    ? `data-ppt-alt-text="${escapeHtml(altText)}"`
    : ''
}

function getPPTElementStrokeDashHTMLAttr(element: PPTElement) {
  const dash = getPPTElementStrokeDash(element)

  return dash
    ? ` data-ppt-stroke-dash="${escapeHtml(dash)}"`
    : ''
}

function getPPTElementCornerRadiusHTMLAttr(element: PPTElement) {
  return element.kind === 'shape' && element.shape === 'rect'
    ? ` data-ppt-corner-radius="${escapeHtml(formatPPTShapeCornerRadius(getPPTShapeCornerRadius(element)))}"`
    : ''
}

function getPPTElementFillOpacityHTMLAttr(element: PPTElement) {
  if (element.kind !== 'shape') {
    return ''
  }

  return ` ${getPPTElementFillOpacityDataAttr(element.fill)}`
}

function getPPTElementFillOpacitySvgAttr(element: PPTElement) {
  if (element.kind !== 'shape') {
    return ''
  }

  return getPPTElementFillOpacityDataAttr(element.fill)
}

function getPPTElementFillOpacityDataAttr(fill: PPTFill) {
  return `data-ppt-fill-opacity="${escapeHtml(formatPPTFillOpacity(getPPTFillOpacity(fill)))}"`
}

function getPPTElementCornerRadiusSvgAttr(element: PPTElement) {
  return element.kind === 'shape' && element.shape === 'rect'
    ? `data-ppt-corner-radius="${escapeHtml(formatPPTShapeCornerRadius(getPPTShapeCornerRadius(element)))}"`
    : ''
}

function getPPTElementStrokeDashSvgAttr(element: PPTElement) {
  const dash = getPPTElementStrokeDash(element)

  return dash
    ? `data-ppt-stroke-dash="${escapeHtml(dash)}"`
    : ''
}

function getPPTElementHyperlinkHTMLAttr(element: PPTElement) {
  const hyperlink = getPPTElementHyperlink(element)

  return hyperlink
    ? ` data-ppt-hyperlink-url="${escapeHtml(hyperlink.url)}"`
    : ''
}

function getPPTElementHyperlinkSvgAttr(element: PPTElement) {
  const hyperlink = getPPTElementHyperlink(element)

  return hyperlink
    ? `data-ppt-hyperlink-url="${escapeHtml(hyperlink.url)}"`
    : ''
}

function getPPTElementShadowHTMLAttrs(element: PPTElement) {
  const entries = getPPTElementShadowAttrEntries(element)

  return entries.length > 0
    ? ` ${entries.join(' ')}`
    : ''
}

function getPPTElementShadowAttrEntries(element: PPTElement) {
  const shadow = getPPTElementShadow(element)

  if (!shadow) {
    return []
  }

  return [
    'data-ppt-shadow="true"',
    `data-ppt-shadow-angle="${shadow.angle}"`,
    `data-ppt-shadow-blur="${shadow.blur}"`,
    `data-ppt-shadow-color="${escapeHtml(shadow.color)}"`,
    `data-ppt-shadow-distance="${shadow.distance}"`,
    `data-ppt-shadow-opacity="${escapeHtml(formatPPTElementShadowOpacity(shadow.opacity))}"`,
  ]
}

function getPPTElementFilterStyle(element: PPTElement) {
  const filter = getPPTElementFilter(element)

  return filter ? `filter:${filter}` : ''
}

function getPPTElementSVGStyleAttr(element: PPTElement) {
  const style = getPPTElementFilterStyle(element)

  return style ? `style="${escapeHtml(style)}"` : ''
}

function getPPTElementFilter(element: PPTElement) {
  const shadow = getPPTElementShadow(element)
  const filters = [
    shadow
      ? getSlideEditObjectShadowFilter({
          ...shadow,
          enabled: true,
        })
      : '',
    element.kind === 'image' ? getPPTImageAdjustmentsFilter(element) : '',
  ].filter(Boolean)

  return filters.length > 0 ? filters.join(' ') : ''
}

function getPPTImageAdjustmentsFilter(element: PPTImage) {
  const adjustments = element.adjustments

  if (!adjustments) {
    return ''
  }

  return [
    adjustments.grayscale === true ? 'grayscale(1)' : '',
    getPPTImageAdjustmentFilter('brightness', adjustments.brightness),
    getPPTImageAdjustmentFilter('contrast', adjustments.contrast),
  ].filter(Boolean).join(' ')
}

function getPPTImageAdjustmentFilter(name: string, value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) && value !== 1
    ? `${name}(${Math.max(0, value)})`
    : ''
}

function getPPTElementAnimationHTMLAttrs(element: PPTElement) {
  return getPPTElementAnimationAttrEntries(element)
    .map((attr) => ` ${attr}`)
    .join('')
}

function getPPTTextAutoFitAttr(element: PPTElement) {
  const autoFit = getPPTTextAutoFit(element)

  return autoFit ? ` data-ppt-text-autofit="${autoFit}"` : ''
}

function getPPTTextAutoFitSvgAttr(element: PPTElement) {
  const autoFit = getPPTTextAutoFit(element)

  return autoFit ? `data-ppt-text-autofit="${autoFit}"` : ''
}

function getPPTTextVerticalAlignSvgAttr(element: PPTElement) {
  return isPPTElementWithText(element)
    ? `data-ppt-vertical-align="${escapeHtml(getPPTElementTextVerticalAlign(element))}"`
    : ''
}

function getPPTTextInsetSvgAttr(element: PPTElement) {
  return isPPTElementWithText(element)
    ? `data-ppt-text-inset="${escapeHtml(formatPPTTextInsetData(getPPTElementTextInset(element)))}"`
    : ''
}

function getPPTTextAutoFit(element: PPTElement) {
  if (element.kind === 'textBox') {
    return element.textAutoFit
  }

  if (
    (element.kind === 'freeform' || element.kind === 'shape') &&
    element.textBody
  ) {
    return element.textAutoFit
  }

  return undefined
}

function getPPTElementSVGAttrs(element: PPTElement) {
  const transform = getPPTElementSVGTransform(element)

  return [
    `data-ppt-element="${escapeHtml(element.id)}"`,
    `data-ppt-kind="${element.kind}"`,
    getPPTElementAltTextSvgAttr(element),
    getPPTElementHyperlinkSvgAttr(element),
    getPPTElementCornerRadiusSvgAttr(element),
    getPPTElementFillOpacitySvgAttr(element),
    getPPTElementStrokeDashSvgAttr(element),
    `data-ppt-opacity="${escapeHtml(formatPPTElementOpacity(getPPTElementOpacity(element)))}"`,
    `opacity="${escapeHtml(formatPPTElementOpacity(getPPTElementOpacity(element)))}"`,
    ...getPPTElementShadowAttrEntries(element),
    element.kind === 'shape'
      ? `data-ppt-shape="${element.shape}"`
      : '',
    getPPTTextAutoFitSvgAttr(element),
    getPPTTextInsetSvgAttr(element),
    getPPTTextVerticalAlignSvgAttr(element),
    ...getPPTElementAnimationAttrEntries(element),
    element.flipH === true ? 'data-ppt-flip-h="true"' : '',
    element.flipV === true ? 'data-ppt-flip-v="true"' : '',
    getPPTElementSVGStyleAttr(element),
    transform ? `transform="${escapeHtml(transform)}"` : '',
  ].filter(Boolean).join(' ')
}

function getPPTElementAnimationAttrEntries(element: PPTElement) {
  if (!element.animation) {
    return []
  }

  return [
    `data-ppt-animation-type="${escapeHtml(element.animation.type)}"`,
    `data-ppt-animation-trigger="${escapeHtml(element.animation.trigger)}"`,
    `data-ppt-animation-duration="${element.animation.durationMs}"`,
    `data-ppt-animation-delay="${element.animation.delayMs}"`,
    `data-ppt-animation-order="${element.animation.order}"`,
  ]
}

function getPPTElementSVGTransform(element: PPTElement) {
  return createPPTCanvasSvgBoundsTransform({
    bounds: element.geometry,
    flipX: element.flipH === true,
    flipY: element.flipV === true,
    rotation: element.geometry.rotation,
  })
}

function getPPTLinePath(element: PPTLine) {
  const bend = Math.min(0.92, Math.max(0.08, element.routeBend ?? 0.5))
  const bendX = element.start.x + (element.end.x - element.start.x) * bend

  return createPPTCanvasSvgPathData([
    { x: element.start.x, y: element.start.y },
    { x: bendX, y: element.start.y },
    { x: bendX, y: element.end.y },
    { x: element.end.x, y: element.end.y },
  ])
}

function getPPTLineSVGPath(element: PPTLine) {
  const bend = Math.min(0.92, Math.max(0.08, element.routeBend ?? 0.5))
  const bendX = element.start.x + (element.end.x - element.start.x) * bend
  const points = [
    { x: element.start.x, y: element.start.y },
    { x: bendX, y: element.start.y },
    { x: bendX, y: element.end.y },
    { x: element.end.x, y: element.end.y },
  ].map((point) => getPPTLineWorldPoint(element, point))

  return createPPTCanvasSvgPathData(points)
}

function getPPTFreeformWorldPathData(element: PPTFreeform) {
  if (element.segments && element.segments.length > 0) {
    return createPPTCanvasSvgPathSegmentData(element.segments.map((segment) => {
      const point = {
        x: element.geometry.x + segment.point.x,
        y: element.geometry.y + segment.point.y,
      }

      if (segment.type !== 'cubic') {
        return {
          point,
          type: segment.type,
        }
      }

      return {
        control1: {
          x: element.geometry.x + segment.control1.x,
          y: element.geometry.y + segment.control1.y,
        },
        control2: {
          x: element.geometry.x + segment.control2.x,
          y: element.geometry.y + segment.control2.y,
        },
        point,
        type: segment.type,
      }
    }))
  }

  const points = element.points.map((point) => ({
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  }))

  return element.pointMode === 'polyline'
    ? createPPTCanvasSvgPathData(points)
    : createPPTCanvasSvgFreehandPathData(points)
}

function getPPTFreeformPathData(element: PPTFreeform) {
  if (element.segments && element.segments.length > 0) {
    return createPPTCanvasSvgPathSegmentData(element.segments)
  }

  return element.pointMode === 'polyline'
    ? createPPTCanvasSvgPathData(element.points)
    : createPPTCanvasSvgFreehandPathData(element.points)
}

function getPPTLineWorldPoint(line: PPTLine, point: PPTLinePoint) {
  return {
    x: line.geometry.x + point.x,
    y: line.geometry.y + point.y,
  }
}

function exportTextStyle(style: PPTTextStyle) {
  const fontWeight = style.fontWeight === 'bold'
    ? '700'
    : style.fontWeight === 'semibold'
      ? '600'
      : '400'

  return [
    `color:${style.color}`,
    `font-family:${getPPTTextFontFamilyCSS(style.fontFamily)}`,
    `font-size:${style.fontSize}px`,
    `font-weight:${fontWeight}`,
    `line-height:${PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT}`,
  ].join(';')
}

function isPPTElementWithText(element: PPTElement) {
  return element.kind === 'textBox' ||
    ((element.kind === 'freeform' || element.kind === 'shape') && !!element.textBody)
}

function normalizePPTTextFontFamily(fontFamily: string | undefined) {
  return PPT_TEXT_FONT_FAMILY_VALUES.has(fontFamily ?? '')
    ? fontFamily ?? PPT_DEFAULT_TEXT_FONT_FAMILY
    : PPT_DEFAULT_TEXT_FONT_FAMILY
}

function getPPTTextFontFamilyCSS(fontFamily: string | undefined) {
  const normalized = normalizePPTTextFontFamily(fontFamily)

  return PPT_TEXT_FONT_FAMILY_OPTIONS.find((option) => option.value === normalized)?.css ??
    PPT_TEXT_FONT_FAMILY_OPTIONS[0].css
}

function getPPTElementTextVerticalAlign(element: PPTElement) {
  const verticalAlign =
    element.kind === 'freeform' || element.kind === 'shape' || element.kind === 'textBox'
    ? element.style?.verticalAlign
    : undefined

  return normalizePPTTextVerticalAlign(
    verticalAlign,
    element.kind === 'freeform' || element.kind === 'shape'
      ? 'middle'
      : PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
  )
}

function normalizePPTTextVerticalAlign(
  verticalAlign: string | undefined,
  fallback: PPTTextVerticalAlign = PPT_DEFAULT_TEXT_VERTICAL_ALIGN,
) {
  const normalizedFallback = normalizeSlideEditTextVerticalAlignment(fallback)
  const normalizedValue = normalizeSlideEditTextVerticalAlignment(verticalAlign)

  return verticalAlign === undefined ? normalizedFallback : normalizedValue
}

function getPPTTextVerticalAlignCSS(verticalAlign: string | undefined) {
  return getSlideEditTextVerticalAlignmentFlexAlignItems(
    normalizePPTTextVerticalAlign(verticalAlign),
  )
}

function getPPTElementTextInset(element: PPTElement): PPTTextInset {
  const fallback = element.kind === 'freeform' || element.kind === 'shape'
    ? PPT_DEFAULT_SHAPE_TEXT_INSET
    : PPT_DEFAULT_TEXT_BOX_INSET
  const inset =
    element.kind === 'freeform' || element.kind === 'shape' || element.kind === 'textBox'
    ? element.style?.textInset
    : undefined

  return {
    bottom: normalizePPTTextInset(inset?.bottom ?? fallback.bottom),
    left: normalizePPTTextInset(inset?.left ?? fallback.left),
    right: normalizePPTTextInset(inset?.right ?? fallback.right),
    top: normalizePPTTextInset(inset?.top ?? fallback.top),
  }
}

function normalizePPTTextInset(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 0

  return normalizeSlideEditTextFrameInsetValue(finiteValue)
}

function getPPTTextInsetCSS(inset: PPTTextInset) {
  return getSlideEditTextFrameInsetPaddingCSS(inset)
}

function formatPPTTextInsetData(inset: PPTTextInset) {
  return `${inset.top},${inset.right},${inset.bottom},${inset.left}`
}

function getPPTElementOpacity(element: PPTElement) {
  return normalizePPTElementOpacity(element.opacity ?? 1)
}

function normalizePPTElementOpacity(value: number) {
  return normalizeSlideEditObjectOpacity(value)
}

function formatPPTElementOpacity(value: number) {
  return toSlideEditObjectOpacityAttributeValue(value)
}

function getPPTElementHyperlink(element: PPTElement) {
  return element.hyperlink ? normalizePPTElementHyperlink(element.hyperlink) : null
}

function getPPTElementAltText(element: PPTElement) {
  return element.accessibility
    ? normalizePPTElementAccessibility(element.accessibility)?.altText
    : undefined
}

function getPPTElementStrokeDash(element: PPTElement) {
  const stroke = getPPTElementStroke(element)

  return stroke ? getPPTStrokeDash(stroke) : undefined
}

function getPPTShapeCornerRadius(element: PPTShape) {
  return element.shape === 'rect'
    ? normalizePPTShapeCornerRadius(element.cornerRadius ?? PPT_SHAPE_CORNER_RADIUS_DEFAULT)
    : 0
}

function normalizePPTShapeCornerRadius(value: number) {
  return normalizeSlideEditObjectCornerRadius(value)
}

function formatPPTShapeCornerRadius(value: number) {
  return toSlideEditObjectCornerRadiusAttributeValue(value)
}

function getPPTFillOpacity(fill: PPTFill) {
  return normalizePPTFillOpacity(fill.opacity ?? 1)
}

function normalizePPTFillOpacity(value: number) {
  return normalizeSlideEditObjectFillOpacity(value)
}

function formatPPTFillOpacity(value: number) {
  return toSlideEditObjectFillOpacityAttributeValue(value)
}

function getPPTFillOpacitySvgAttr(fill: PPTFill) {
  const opacity = getPPTFillOpacity(fill)

  return opacity === 1
    ? ''
    : ` fill-opacity="${escapeHtml(formatPPTFillOpacity(opacity))}"`
}

function getPPTFillColorCSS(fill: PPTFill) {
  const opacity = getPPTFillOpacity(fill)

  if (opacity === 1) {
    return fill.color
  }

  return getSlideEditColorWithAlphaCSS({
    color: fill.color,
    opacity,
  })
}

function getPPTElementStroke(element: PPTElement): PPTStroke | null {
  if (element.kind === 'shape' || element.kind === 'image') {
    return element.stroke ? normalizePPTStroke(element.stroke) : null
  }

  if (element.kind === 'line' || element.kind === 'freeform') {
    return normalizePPTStroke(element.stroke)
  }

  return null
}

function normalizePPTStroke(stroke: Partial<PPTStroke>): PPTStroke {
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

function normalizePPTStrokeWidth(value: number) {
  const finiteValue = Number.isFinite(value) ? value : 2

  return Math.max(0, Math.min(40, finiteValue))
}

function getPPTStrokeDash(stroke: PPTStroke | undefined): PPTStrokeDash {
  return normalizePPTStrokeDash(stroke?.dash)
}

function normalizePPTStrokeDash(value: unknown): PPTStrokeDash {
  return normalizeSlideEditObjectStrokeLineStyle(
    typeof value === 'string' ? value : null,
  ) as PPTStrokeDash
}

function getPPTStrokeDashBorderStyle(stroke: PPTStroke | undefined) {
  return getSlideEditObjectStrokeLineStyleBorderStyle(getPPTStrokeDash(stroke))
}

function getPPTStrokeDashArray(stroke: PPTStroke | undefined) {
  const dashArray = getSlideEditObjectStrokeLineStyleDashArray({
    strokeWidth: normalizePPTStrokeWidth(stroke?.width ?? 2),
    value: getPPTStrokeDash(stroke),
  })

  return dashArray
    ?.split(' ')
    .map((value) => formatNumber(Number(value)))
    .join(' ') ?? ''
}

function getPPTStrokeDashArraySvgAttr(stroke: PPTStroke | undefined) {
  const dashArray = getPPTStrokeDashArray(stroke)

  return dashArray ? ` stroke-dasharray="${escapeHtml(dashArray)}"` : ''
}

function getPPTImageAltText(element: PPTImage) {
  return getPPTElementAltText(element) ?? element.alt
}

function normalizePPTElementAccessibility(
  accessibility: Partial<PPTElementAccessibility>,
): PPTElementAccessibility | null {
  const altText = normalizePPTAltText(accessibility.altText ?? '')

  return altText ? { altText } : null
}

function normalizePPTAltText(value: string) {
  const normalized = value.trim().slice(0, PPT_ALT_TEXT_MAX_LENGTH)

  if (!normalized || hasPPTControlCharacter(normalized)) {
    return ''
  }

  return normalized
}

function normalizePPTElementHyperlink(
  hyperlink: Partial<PPTElementHyperlink>,
): PPTElementHyperlink | null {
  const url = normalizePPTElementHyperlinkUrl(hyperlink.url ?? '')

  return url ? { url } : null
}

function normalizePPTElementHyperlinkUrl(url: string) {
  const normalized = url.trim().slice(0, PPT_HYPERLINK_URL_MAX_LENGTH)

  if (!normalized || !isPPTElementHyperlinkUrlAllowed(normalized)) {
    return ''
  }

  return normalized
}

function isPPTElementHyperlinkUrlAllowed(url: string) {
  return !hasPPTControlCharacter(url) &&
    !/^(javascript|data|vbscript):/i.test(url)
}

function hasPPTControlCharacter(value: string) {
  return [...value].some((char) => {
    const code = char.charCodeAt(0)

    return code <= 31 || code === 127
  })
}

function getPPTElementShadow(element: PPTElement): PPTElementShadow | null {
  return element.shadow ? normalizePPTElementShadow(element.shadow) : null
}

function normalizePPTElementShadow(
  shadow: Partial<PPTElementShadow>,
): PPTElementShadow {
  return {
    angle: normalizePPTElementShadowAngle(shadow.angle ?? PPT_DEFAULT_ELEMENT_SHADOW.angle),
    blur: normalizePPTElementShadowBlur(shadow.blur ?? PPT_DEFAULT_ELEMENT_SHADOW.blur),
    color: normalizePPTElementShadowColor(shadow.color ?? PPT_DEFAULT_ELEMENT_SHADOW.color),
    distance: normalizePPTElementShadowDistance(shadow.distance ?? PPT_DEFAULT_ELEMENT_SHADOW.distance),
    opacity: normalizePPTElementShadowOpacity(shadow.opacity ?? PPT_DEFAULT_ELEMENT_SHADOW.opacity),
  }
}

function normalizePPTElementShadowAngle(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.angle

  return Math.min(
    PPT_ELEMENT_SHADOW_ANGLE_MAX,
    Math.max(PPT_ELEMENT_SHADOW_ANGLE_MIN, Math.round(finiteValue)),
  )
}

function normalizePPTElementShadowBlur(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.blur

  return Math.min(PPT_ELEMENT_SHADOW_BLUR_MAX, Math.max(0, Math.round(finiteValue)))
}

function normalizePPTElementShadowDistance(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.distance

  return Math.min(PPT_ELEMENT_SHADOW_DISTANCE_MAX, Math.max(0, Math.round(finiteValue)))
}

function normalizePPTElementShadowOpacity(value: number) {
  const finiteValue = Number.isFinite(value) ? value : PPT_DEFAULT_ELEMENT_SHADOW.opacity
  const clamped = Math.min(
    PPT_ELEMENT_SHADOW_OPACITY_MAX,
    Math.max(PPT_ELEMENT_SHADOW_OPACITY_MIN, finiteValue),
  )

  return Math.round(clamped * 100) / 100
}

function formatPPTElementShadowOpacity(value: number) {
  return String(normalizePPTElementShadowOpacity(value))
}

function normalizePPTElementShadowColor(color: string) {
  return /^#[\da-f]{6}$/i.test(color) ? color : PPT_DEFAULT_ELEMENT_SHADOW.color
}

function getPPTTextVerticalAlignOffset({
  body,
  fontSize,
  geometry,
  inset,
  verticalAlign,
}: {
  body: PPTTextBody
  fontSize: number
  geometry: PPTElement['geometry']
  inset: PPTTextInset
  verticalAlign: PPTTextVerticalAlign
}) {
  if (verticalAlign === 'top') {
    return 0
  }

  const innerHeight = Math.max(0, geometry.h - inset.top - inset.bottom)
  const textHeight = getPPTTextBodySVGHeight(body, fontSize)
  const available = Math.max(0, innerHeight - textHeight)

  return verticalAlign === 'middle' ? available / 2 : available
}

function getPPTTextBodySVGHeight(body: PPTTextBody, fontSize: number) {
  return body.paragraphs.reduce(
    (height, paragraph) =>
      height +
      getPPTParagraphSpacingBefore(paragraph) +
      fontSize * getPPTParagraphLineHeight(paragraph) +
      getPPTParagraphSpacingAfter(paragraph),
    0,
  )
}

function getPPTParagraphHTMLAttrs(paragraph: PPTParagraph) {
  return [
    ` data-ppt-line-height="${formatNumber(getPPTParagraphLineHeight(paragraph))}"`,
    ` data-ppt-list-level="${formatNumber(getPPTParagraphListLevel(paragraph))}"`,
    ` data-ppt-spacing-after="${formatNumber(getPPTParagraphSpacingAfter(paragraph))}"`,
    ` data-ppt-spacing-before="${formatNumber(getPPTParagraphSpacingBefore(paragraph))}"`,
  ].join('')
}

function getPPTParagraphStyleAttr(paragraph: PPTParagraph) {
  return ` style="${[
    `line-height:${formatNumber(getPPTParagraphLineHeight(paragraph))}`,
    `margin-bottom:${formatNumber(getPPTParagraphSpacingAfter(paragraph))}px`,
    `margin-top:${formatNumber(getPPTParagraphSpacingBefore(paragraph))}px`,
    `--ppt-paragraph-list-level-indent:${
      getSlideEditTextParagraphListLevelIndentCSSValue(
        getPPTParagraphListLevel(paragraph),
      )
    }`,
  ].join(';')}"`
}

function getPPTParagraphListLevel(paragraph: PPTParagraph) {
  return normalizePPTParagraphListLevel(paragraph.level)
}

function getPPTParagraphLineHeight(paragraph: PPTParagraph) {
  return normalizePPTParagraphLineHeight(
    paragraph.lineHeight ?? PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
  )
}

function getPPTParagraphSpacingAfter(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingAfter ?? 0)
}

function getPPTParagraphSpacingBefore(paragraph: PPTParagraph) {
  return normalizePPTParagraphSpacing(paragraph.spacingBefore ?? 0)
}

function normalizePPTParagraphListLevel(value: number | null | undefined) {
  return normalizeSlideEditTextParagraphListLevel(value)
}

function normalizePPTParagraphLineHeight(value: number) {
  const next = Math.min(
    PPT_PARAGRAPH_LINE_HEIGHT_MAX,
    Math.max(
      PPT_PARAGRAPH_LINE_HEIGHT_MIN,
      Number.isFinite(value) ? value : PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT,
    ),
  )

  return Math.round(next * 100) / 100
}

function normalizePPTParagraphSpacing(value: number) {
  return Math.min(
    PPT_PARAGRAPH_SPACING_MAX,
    Math.max(0, Number.isFinite(value) ? Math.round(value) : 0),
  )
}

function exportShapeStyle(element: PPTShape) {
  return [
    `background:${getPPTFillColorCSS(element.fill)}`,
    element.shape === 'rect'
      ? `border-radius:${formatPPTShapeCornerRadius(getPPTShapeCornerRadius(element))}px`
      : '',
    element.stroke
      ? `border:${element.stroke.width}px solid ${element.stroke.color}`
      : '',
    element.stroke
      ? `border-style:${getPPTStrokeDashBorderStyle(element.stroke)}`
      : '',
  ].filter(Boolean).join(';')
}

function toPercent(value: number, total: number) {
  return `${Number(((value / total) * 100).toFixed(4))}%`
}

function getVisiblePPTElements(elements: PPTElement[]) {
  return elements.filter((element) => element.visible !== false)
}

function getPPTElementsExportBounds(elements: PPTElement[]) {
  return unionPPTCanvasRectList(elements.map((element) => element.geometry))
}

function getPPTSvgImageAlignX(value: number) {
  if (value <= 33) {
    return 'xMin'
  }

  if (value >= 67) {
    return 'xMax'
  }

  return 'xMid'
}

function getPPTSvgImageAlignY(value: number) {
  if (value <= 33) {
    return 'YMin'
  }

  if (value >= 67) {
    return 'YMax'
  }

  return 'YMid'
}

function getPPTSvgTextX({
  align,
  geometry,
  inset,
}: {
  align: PPTParagraph['align']
  geometry: PPTElement['geometry']
  inset: PPTTextInset
}) {
  if (align === 'center') {
    return geometry.x + inset.left + (geometry.w - inset.left - inset.right) / 2
  }

  if (align === 'right') {
    return geometry.x + geometry.w - inset.right
  }

  return geometry.x + inset.left
}

function getPPTSvgTextAnchor(align: PPTParagraph['align']) {
  if (align === 'center') {
    return 'middle'
  }

  if (align === 'right') {
    return 'end'
  }

  return 'start'
}

function getPPTSvgFontWeight(style: PPTTextStyle | undefined) {
  if (style?.fontWeight === 'bold') {
    return '700'
  }

  if (style?.fontWeight === 'semibold') {
    return '600'
  }

  return '400'
}

function getPPTTableColumnCount(rows: readonly (readonly string[])[]) {
  return Math.max(0, ...rows.map((row) => row.length))
}

function formatNumber(value: number) {
  return formatPPTCanvasSvgNumber(value)
}

function escapeHtml(value: string) {
  return escapePPTCanvasXmlAttribute(value)
}

function escapeScriptJson(value: string) {
  return value
    .replaceAll('&', '\\u0026')
    .replaceAll('<', '\\u003c')
    .replaceAll('>', '\\u003e')
}

function indent(value: string, spaces: number) {
  const prefix = ' '.repeat(spaces)

  return value.split('\n').map((line) => `${prefix}${line}`).join('\n')
}
