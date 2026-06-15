import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTDeck,
  type PPTElement,
  type PPTLine,
  type PPTRun,
  type PPTShape,
  type PPTTextBody,
  type PPTTextStyle,
} from './pptModel'

export function exportPPTDeckHTML(deck: PPTDeck) {
  const body = deck.slides.map((slide) => {
    const elements = slide.elements
      .filter((element) => element.visible !== false)
      .map(renderPPTElementHTML)
      .join('\n')

    return [
      `  <section class="ppt-slide" data-ppt-slide="${escapeHtml(slide.id)}" style="background:${escapeHtml(slide.background?.color ?? '#ffffff')}">`,
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

  const text = renderPPTTextBodyHTML(element.textBody)
  const textStyle = element.style ? exportTextStyle(element.style) : ''
  const paragraphStyle = `text-align:${element.textBody?.paragraphs[0]?.align ?? 'left'}`
  const bulletListAttr = element.textBody && hasPPTTextBodyBullet(element.textBody)
    ? ' data-ppt-bullet-list="true"'
    : ''

  if (element.kind === 'shape') {
    return `    <div class="ppt-element ppt-shape ppt-shape-${element.shape}" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${bulletListAttr} style="${[...style, exportShapeStyle(element), textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
  }

  return `    <div class="ppt-element ppt-text" data-ppt-element="${escapeHtml(element.id)}"${transformAttrs}${bulletListAttr} style="${[...style, textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
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
    '.ppt-line{display:block;overflow:visible;padding:0;}',
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

    return `<span class="ppt-text-paragraph"${bulletAttr}>${runs}</span>`
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
  ].join('')
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
    'line-height:1.14',
  ].join(';')
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
