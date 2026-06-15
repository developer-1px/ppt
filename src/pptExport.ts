import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTComment,
  type PPTDeck,
  type PPTElement,
  type PPTFreeform,
  type PPTImage,
  type PPTLine,
  type PPTLinePoint,
  type PPTParagraph,
  type PPTRun,
  type PPTShape,
  type PPTSlide,
  type PPTTable,
  type PPTTextBody,
  type PPTTextStyle,
} from './pptModel'

const PPT_SELECTION_EXPORT_PADDING = 24
const PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT = 1.14
const PPT_PARAGRAPH_LINE_HEIGHT_MIN = 0.8
const PPT_PARAGRAPH_LINE_HEIGHT_MAX = 3
const PPT_PARAGRAPH_SPACING_MAX = 240

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
    const transitionAttrs = getPPTSlideTransitionAttrs(slide, 'data-ppt-transition')

    return [
      `  <section class="ppt-slide" data-ppt-slide="${escapeHtml(slide.id)}"${layoutAttr}${themeAttr}${transitionAttrs} style="background:${escapeHtml(slide.background?.color ?? '#ffffff')}">`,
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
    getPPTSlideTransitionAttrs(slide, 'data-ppt-svg-transition').trim(),
  ].filter(Boolean)

  return attrs.length > 0 ? ` ${attrs.join(' ')}` : ''
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
  const style = [
    `left:${toPercent(element.geometry.x, PPT_SLIDE_WIDTH)}`,
    `top:${toPercent(element.geometry.y, PPT_SLIDE_HEIGHT)}`,
    `width:${toPercent(element.geometry.w, PPT_SLIDE_WIDTH)}`,
    `height:${toPercent(element.geometry.h, PPT_SLIDE_HEIGHT)}`,
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

    return `    <img class="ppt-element ppt-image" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs} data-ppt-image-fit="${fit}" data-ppt-image-crop-x="${crop.x}" data-ppt-image-crop-y="${crop.y}" alt="${escapeHtml(element.alt)}" src="${escapeHtml(element.src)}" style="${[...style, `object-fit:${fit}`, `object-position:${crop.x}% ${crop.y}%`].filter(Boolean).join(';')}" />`
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
  const bulletListAttr = element.textBody && hasPPTTextBodyBullet(element.textBody)
    ? ' data-ppt-bullet-list="true"'
    : ''
  const autoFitAttr = getPPTTextAutoFitAttr(element)

  if (element.kind === 'shape') {
    return `    <div class="ppt-element ppt-shape ppt-shape-${element.shape}" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${bulletListAttr}${autoFitAttr} style="${[...style, exportShapeStyle(element), textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
  }

  return `    <div class="ppt-element ppt-text" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${bulletListAttr}${autoFitAttr} style="${[...style, textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
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
    inset: element.kind === 'shape' ? 18 : 0,
    style: element.style,
  })

  if (element.kind === 'shape') {
    return `<g ${attrs}>${renderPPTShapeSVG(element)}${text}</g>`
  }

  return `<g ${attrs}>${text}</g>`
}

function renderPPTImageSVG(element: PPTImage) {
  const crop = element.crop ?? { x: 50, y: 50 }
  const fit = element.fit ?? 'cover'
  const preserveAspectRatio = `${getPPTSvgImageAlignX(crop.x)}${getPPTSvgImageAlignY(crop.y)} ${fit === 'contain' ? 'meet' : 'slice'}`
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-image-fit="${fit}"`,
    `data-ppt-image-crop-x="${formatNumber(crop.x)}"`,
    `data-ppt-image-crop-y="${formatNumber(crop.y)}"`,
  ].join(' ')

  return `<g ${attrs}><image href="${escapeHtml(element.src)}" x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" preserveAspectRatio="${preserveAspectRatio}"><title>${escapeHtml(element.alt)}</title></image></g>`
}

function renderPPTShapeSVG(element: PPTShape) {
  const fill = element.fill.color
  const stroke = element.stroke
    ? ` stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}"`
    : ' stroke="none"'

  if (element.shape === 'ellipse') {
    return `<ellipse cx="${formatNumber(element.geometry.x + element.geometry.w / 2)}" cy="${formatNumber(element.geometry.y + element.geometry.h / 2)}" rx="${formatNumber(element.geometry.w / 2)}" ry="${formatNumber(element.geometry.h / 2)}" fill="${escapeHtml(fill)}"${stroke} />`
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

    return `<polygon points="${points}" fill="${escapeHtml(fill)}"${stroke} />`
  }

  return `<rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" rx="24" fill="${escapeHtml(fill)}"${stroke} />`
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
    ? `<path data-ppt-line-path d="${getPPTLineSVGPath(element)}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}" stroke-linecap="round" stroke-linejoin="round"${markerStart}${markerEnd}></path>`
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

  return `<line x1="${formatNumber(start.x)}" y1="${formatNumber(start.y)}" x2="${formatNumber(end.x)}" y2="${formatNumber(end.y)}" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}" stroke-linecap="round"${markerStart}${markerEnd}></line>`
}

function renderPPTTextBodySVG({
  body,
  geometry,
  inset,
  style,
}: {
  body: PPTTextBody | undefined
  geometry: PPTElement['geometry']
  inset: number
  style: PPTTextStyle | undefined
}) {
  if (!body) {
    return ''
  }

  const fontSize = style?.fontSize ?? 24
  let y = geometry.y + inset

  return body.paragraphs.map((paragraph) => {
    const bullet = paragraph.bullet === 'bullet'
    const runs = paragraph.runs.map(renderPPTTextRunSVG).join('')
    const bulletPrefix = bullet ? '<tspan data-ppt-bullet="true">&#8226; </tspan>' : ''
    const lineHeight = getPPTParagraphLineHeight(paragraph)
    const spacingBefore = getPPTParagraphSpacingBefore(paragraph)
    const spacingAfter = getPPTParagraphSpacingAfter(paragraph)
    const textAnchor = getPPTSvgTextAnchor(paragraph.align)
    const x = getPPTSvgTextX({
      align: paragraph.align,
      geometry,
      inset,
    })
    y += spacingBefore + fontSize
    const attrs = [
      'class="ppt-svg-text-paragraph"',
      bullet ? 'data-ppt-bullet="true"' : '',
      `data-ppt-line-height="${formatNumber(lineHeight)}"`,
      `data-ppt-spacing-after="${formatNumber(spacingAfter)}"`,
      `data-ppt-spacing-before="${formatNumber(spacingBefore)}"`,
      `x="${formatNumber(x)}"`,
      `y="${formatNumber(y)}"`,
      `fill="${escapeHtml(style?.color ?? '#111827')}"`,
      'font-family="Inter, Arial, sans-serif"',
      `font-size="${formatNumber(fontSize)}"`,
      `font-weight="${getPPTSvgFontWeight(style)}"`,
      `text-anchor="${textAnchor}"`,
    ].filter(Boolean).join(' ')

    y += fontSize * lineHeight + spacingAfter

    return `<text ${attrs}>${bulletPrefix}${runs}</text>`
  }).join('')
}

function renderPPTTextRunSVG(run: PPTRun) {
  const attrs = [
    run.italic === true ? 'data-ppt-run-italic="true" font-style="italic"' : '',
    run.underline === true ? 'data-ppt-run-underline="true" text-decoration="underline"' : '',
    run.bold === true ? 'font-weight="700"' : '',
    run.color ? `fill="${escapeHtml(run.color)}"` : '',
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
    '.ppt-text-paragraph{display:block;min-height:1em;}',
    '.ppt-text-paragraph[data-ppt-bullet="true"]{position:relative;padding-left:1.1em;}',
    '.ppt-text-paragraph[data-ppt-bullet="true"]::before{content:"\\2022";position:absolute;left:0;}',
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
    const bulletAttr = paragraph.bullet === 'bullet'
      ? ' data-ppt-bullet="true"'
      : ''
    const paragraphAttrs = getPPTParagraphHTMLAttrs(paragraph)
    const paragraphStyleAttr = getPPTParagraphStyleAttr(paragraph)

    return `<span class="ppt-text-paragraph"${bulletAttr}${paragraphAttrs}${paragraphStyleAttr}>${runs}</span>`
  }).join('')
}

function renderPPTTextRunHTML(run: PPTRun) {
  const attrs = [
    run.italic === true ? 'data-ppt-run-italic="true"' : '',
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
    run.color ? `color:${escapeHtml(run.color)}` : '',
    run.italic === true ? 'font-style:italic' : '',
    run.size ? `font-size:${run.size}px` : '',
    run.underline === true ? 'text-decoration:underline' : '',
  ].filter(Boolean).join(';')

  return styles ? `style="${styles}"` : ''
}

function hasPPTTextBodyBullet(body: PPTTextBody) {
  return body.paragraphs.some((paragraph) => paragraph.bullet === 'bullet')
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
    ? `<path data-ppt-line-path d="${getPPTLinePath(element)}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${element.stroke.width}" stroke-linecap="round" stroke-linejoin="round"${markerStart}${markerEnd}></path>`
    : `<line x1="${element.start.x}" y1="${element.start.y}" x2="${element.end.x}" y2="${element.end.y}" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${element.stroke.width}" stroke-linecap="round"${markerStart}${markerEnd}></line>`

  const connectionAttrs = [
    `data-ppt-line-route="${route}"`,
    element.startConnection
      ? `data-ppt-start-connection="${escapeHtml(element.startConnection.elementId)}:${element.startConnection.anchor}"`
      : '',
    element.endConnection
      ? `data-ppt-end-connection="${escapeHtml(element.endConnection.elementId)}:${element.endConnection.anchor}"`
      : '',
  ].filter(Boolean).join(' ')

  return `    <svg class="ppt-element ppt-line" data-ppt-element="${escapeHtml(element.id)}" ${connectionAttrs} style="${style.filter(Boolean).join(';')}" viewBox="0 0 ${element.geometry.w} ${element.geometry.h}" preserveAspectRatio="none" aria-hidden="true">${marker}${lineMarkup}</svg>`
}

function renderPPTFreeformHTML(element: PPTFreeform, style: string[]) {
  const attrs = [
    `data-ppt-element="${escapeHtml(element.id)}"`,
    `data-ppt-kind="freeform"`,
    `data-ppt-freeform-points="${element.points.length}"`,
  ].join(' ')

  return `    <svg class="ppt-element ppt-freeform" ${attrs} style="${style.filter(Boolean).join(';')}" viewBox="0 0 ${formatNumber(element.geometry.w)} ${formatNumber(element.geometry.h)}" preserveAspectRatio="none" aria-hidden="true"><path data-ppt-freeform-path d="${escapeHtml(getPPTFreeformPathData(element.points))}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}" stroke-linecap="round" stroke-linejoin="round" opacity="${formatNumber(element.opacity ?? 1)}"></path></svg>`
}

function renderPPTFreeformSVG(element: PPTFreeform) {
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-freeform-points="${element.points.length}"`,
  ].join(' ')

  return `<g ${attrs}><path data-ppt-freeform-path d="${escapeHtml(getPPTFreeformWorldPathData(element))}" fill="none" stroke="${escapeHtml(element.stroke.color)}" stroke-width="${formatNumber(element.stroke.width)}" stroke-linecap="round" stroke-linejoin="round" opacity="${formatNumber(element.opacity ?? 1)}"></path></g>`
}

function renderPPTCommentHTML(
  element: PPTComment,
  style: string[],
  transformAttrs: string,
) {
  const author = element.authorName ?? 'You'
  const createdAt = element.createdAt ?? 'Just now'
  const resolvedAttr = element.resolved === true
    ? ' data-ppt-comment-resolved="true"'
    : ''

  return `    <div class="ppt-element ppt-comment" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${resolvedAttr} style="${style.filter(Boolean).join(';')}"><div class="ppt-comment-meta"><span>${escapeHtml(author)}</span><span class="ppt-comment-created">${escapeHtml(createdAt)}</span></div><p class="ppt-comment-body">${escapeHtml(element.body)}</p></div>`
}

function renderPPTCommentSVG(element: PPTComment) {
  const attrs = [
    getPPTElementSVGAttrs(element),
    element.resolved === true ? 'data-ppt-comment-resolved="true"' : '',
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

function renderPPTTableHTML(
  element: PPTTable,
  style: string[],
  transformAttrs: string,
) {
  const columnCount = getPPTTableColumnCount(element.rows)
  const rowCount = element.rows.length
  const head = element.rows[0]
    ? `<thead><tr>${renderPPTTableHTMLRow(element.rows[0], columnCount, 'th')}</tr></thead>`
    : ''
  const bodyRows = element.rows
    .slice(1)
    .map((row) => `<tr>${renderPPTTableHTMLRow(row, columnCount, 'td')}</tr>`)
    .join('')
  const body = `<tbody>${bodyRows}</tbody>`
  const attrs = [
    `class="ppt-element ppt-table"`,
    `data-ppt-element="${escapeHtml(element.id)}"`,
    transformAttrs.trim(),
    `data-ppt-table-rows="${rowCount}"`,
    `data-ppt-table-cols="${columnCount}"`,
    `style="${style.filter(Boolean).join(';')}"`,
  ].filter(Boolean).join(' ')

  return `    <table ${attrs}>${head}${body}</table>`
}

function renderPPTTableHTMLRow(
  row: readonly string[],
  columnCount: number,
  tagName: 'td' | 'th',
) {
  return Array.from({ length: columnCount }, (_, index) =>
    `<${tagName} data-ppt-table-cell="${index}">${escapeHtml(row[index] ?? '')}</${tagName}>`,
  ).join('')
}

function renderPPTTableSVG(element: PPTTable) {
  const columnCount = getPPTTableColumnCount(element.rows)
  const rowCount = element.rows.length
  const cellWidth = columnCount > 0 ? element.geometry.w / columnCount : element.geometry.w
  const cellHeight = rowCount > 0 ? element.geometry.h / rowCount : element.geometry.h
  const fontSize = Math.max(10, Math.min(18, cellHeight * 0.38))
  const cells = element.rows.flatMap((row, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) => {
      const x = element.geometry.x + cellWidth * columnIndex
      const y = element.geometry.y + cellHeight * rowIndex
      const textX = x + Math.min(10, cellWidth * 0.12)
      const textY = y + cellHeight / 2
      const headerAttrs = rowIndex === 0
        ? ' data-ppt-table-header="true" font-weight="700"'
        : ''

      return [
        `<rect data-ppt-table-cell="${rowIndex}:${columnIndex}" x="${formatNumber(x)}" y="${formatNumber(y)}" width="${formatNumber(cellWidth)}" height="${formatNumber(cellHeight)}" fill="${rowIndex === 0 ? '#eff6ff' : '#ffffff'}" stroke="#dbe3ef" stroke-width="1"></rect>`,
        `<text data-ppt-table-text="${rowIndex}:${columnIndex}"${headerAttrs} x="${formatNumber(textX)}" y="${formatNumber(textY)}" fill="#111827" font-family="Inter, Arial, sans-serif" font-size="${formatNumber(fontSize)}" dominant-baseline="middle">${escapeHtml(row[columnIndex] ?? '')}</text>`,
      ].join('')
    }),
  ).join('')
  const attrs = [
    getPPTElementSVGAttrs(element),
    `data-ppt-table-rows="${rowCount}"`,
    `data-ppt-table-cols="${columnCount}"`,
  ].join(' ')

  return `<g ${attrs}><rect x="${formatNumber(element.geometry.x)}" y="${formatNumber(element.geometry.y)}" width="${formatNumber(element.geometry.w)}" height="${formatNumber(element.geometry.h)}" fill="#ffffff" stroke="#94a3b8" stroke-width="1"></rect>${cells}</g>`
}

function getPPTElementTransform(element: PPTElement) {
  const transforms = [
    element.geometry.rotation ? `rotate(${element.geometry.rotation}deg)` : '',
    element.flipH === true ? 'scaleX(-1)' : '',
    element.flipV === true ? 'scaleY(-1)' : '',
  ].filter(Boolean)

  return transforms.length > 0 ? transforms.join(' ') : ''
}

function getPPTElementTransformAttrs(element: PPTElement) {
  return [
    element.flipH === true ? ' data-ppt-flip-h="true"' : '',
    element.flipV === true ? ' data-ppt-flip-v="true"' : '',
    getPPTElementAnimationHTMLAttrs(element),
  ].join('')
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

function getPPTTextAutoFit(element: PPTElement) {
  if (element.kind === 'textBox') {
    return element.textAutoFit
  }

  if (element.kind === 'shape' && element.textBody) {
    return element.textAutoFit
  }

  return undefined
}

function getPPTElementSVGAttrs(element: PPTElement) {
  const transform = getPPTElementSVGTransform(element)

  return [
    `data-ppt-element="${escapeHtml(element.id)}"`,
    `data-ppt-kind="${element.kind}"`,
    element.kind === 'shape'
      ? `data-ppt-shape="${element.shape}"`
      : '',
    getPPTTextAutoFitSvgAttr(element),
    ...getPPTElementAnimationAttrEntries(element),
    element.flipH === true ? 'data-ppt-flip-h="true"' : '',
    element.flipV === true ? 'data-ppt-flip-v="true"' : '',
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
  const scaleX = element.flipH === true ? -1 : 1
  const scaleY = element.flipV === true ? -1 : 1
  const rotation = element.geometry.rotation ?? 0

  if (scaleX === 1 && scaleY === 1 && !rotation) {
    return ''
  }

  const centerX = element.geometry.x + element.geometry.w / 2
  const centerY = element.geometry.y + element.geometry.h / 2
  const transforms = [
    `translate(${formatNumber(centerX)} ${formatNumber(centerY)})`,
    rotation ? `rotate(${formatNumber(rotation)})` : '',
    scaleX !== 1 || scaleY !== 1 ? `scale(${scaleX} ${scaleY})` : '',
    `translate(${formatNumber(-centerX)} ${formatNumber(-centerY)})`,
  ].filter(Boolean)

  return transforms.join(' ')
}

function getPPTLinePath(element: PPTLine) {
  const bend = Math.min(0.92, Math.max(0.08, element.routeBend ?? 0.5))
  const bendX = element.start.x + (element.end.x - element.start.x) * bend

  return [
    `M ${element.start.x} ${element.start.y}`,
    `L ${bendX} ${element.start.y}`,
    `L ${bendX} ${element.end.y}`,
    `L ${element.end.x} ${element.end.y}`,
  ].join(' ')
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

  return [
    `M ${formatNumber(points[0].x)} ${formatNumber(points[0].y)}`,
    `L ${formatNumber(points[1].x)} ${formatNumber(points[1].y)}`,
    `L ${formatNumber(points[2].x)} ${formatNumber(points[2].y)}`,
    `L ${formatNumber(points[3].x)} ${formatNumber(points[3].y)}`,
  ].join(' ')
}

function getPPTFreeformWorldPathData(element: PPTFreeform) {
  return getPPTFreeformPathData(element.points.map((point) => ({
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  })))
}

function getPPTFreeformPathData(points: readonly PPTLinePoint[]) {
  const [first, second, ...rest] = points

  if (!first) {
    return ''
  }

  if (!second) {
    return `M ${formatNumber(first.x)} ${formatNumber(first.y)}`
  }

  if (rest.length === 0) {
    return [
      `M ${formatNumber(first.x)} ${formatNumber(first.y)}`,
      `L ${formatNumber(second.x)} ${formatNumber(second.y)}`,
    ].join(' ')
  }

  return [
    `M ${formatNumber(first.x)} ${formatNumber(first.y)}`,
    `Q ${formatNumber(second.x)} ${formatNumber(second.y)} ${getPPTPathMidpoint(second, rest[0])}`,
    ...rest.slice(1).map((point, index) => {
      const control = rest[index]

      return `Q ${formatNumber(control.x)} ${formatNumber(control.y)} ${getPPTPathMidpoint(control, point)}`
    }),
    `L ${formatNumber(rest[rest.length - 1].x)} ${formatNumber(rest[rest.length - 1].y)}`,
  ].join(' ')
}

function getPPTPathMidpoint(a: PPTLinePoint, b: PPTLinePoint) {
  return `${formatNumber((a.x + b.x) / 2)} ${formatNumber((a.y + b.y) / 2)}`
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
    `font-size:${style.fontSize}px`,
    `font-weight:${fontWeight}`,
    `line-height:${PPT_PARAGRAPH_LINE_HEIGHT_DEFAULT}`,
  ].join(';')
}

function getPPTParagraphHTMLAttrs(paragraph: PPTParagraph) {
  return [
    ` data-ppt-line-height="${formatNumber(getPPTParagraphLineHeight(paragraph))}"`,
    ` data-ppt-spacing-after="${formatNumber(getPPTParagraphSpacingAfter(paragraph))}"`,
    ` data-ppt-spacing-before="${formatNumber(getPPTParagraphSpacingBefore(paragraph))}"`,
  ].join('')
}

function getPPTParagraphStyleAttr(paragraph: PPTParagraph) {
  return ` style="${[
    `line-height:${formatNumber(getPPTParagraphLineHeight(paragraph))}`,
    `margin-bottom:${formatNumber(getPPTParagraphSpacingAfter(paragraph))}px`,
    `margin-top:${formatNumber(getPPTParagraphSpacingBefore(paragraph))}px`,
  ].join(';')}"`
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
    `background:${element.fill.color}`,
    element.stroke
      ? `border:${element.stroke.width}px solid ${element.stroke.color}`
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
  if (elements.length === 0) {
    return null
  }

  const minX = Math.min(...elements.map((element) => element.geometry.x))
  const minY = Math.min(...elements.map((element) => element.geometry.y))
  const maxX = Math.max(...elements.map((element) => element.geometry.x + element.geometry.w))
  const maxY = Math.max(...elements.map((element) => element.geometry.y + element.geometry.h))

  return {
    h: maxY - minY,
    w: maxX - minX,
    x: minX,
    y: minY,
  }
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
  align: 'center' | 'left' | 'right' | undefined
  geometry: PPTElement['geometry']
  inset: number
}) {
  if (align === 'center') {
    return geometry.x + geometry.w / 2
  }

  if (align === 'right') {
    return geometry.x + geometry.w - inset
  }

  return geometry.x + inset
}

function getPPTSvgTextAnchor(align: 'center' | 'left' | 'right' | undefined) {
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
  return Number.isInteger(value) ? `${value}` : `${Number(value.toFixed(3))}`
}

function escapeHtml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
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
