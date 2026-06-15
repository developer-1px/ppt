import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  readPPTText,
  type PPTDeck,
  type PPTElement,
  type PPTShape,
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
  const style = [
    `left:${toPercent(element.geometry.x, PPT_SLIDE_WIDTH)}`,
    `top:${toPercent(element.geometry.y, PPT_SLIDE_HEIGHT)}`,
    `width:${toPercent(element.geometry.w, PPT_SLIDE_WIDTH)}`,
    `height:${toPercent(element.geometry.h, PPT_SLIDE_HEIGHT)}`,
  ]

  if (element.kind === 'image') {
    return `    <img class="ppt-element" data-ppt-element="${escapeHtml(element.id)}" alt="${escapeHtml(element.alt)}" src="${escapeHtml(element.src)}" style="${style.join(';')}" />`
  }

  const text = escapeHtml(readPPTText(element.textBody))
  const textStyle = element.style ? exportTextStyle(element.style) : ''
  const paragraphStyle = `text-align:${element.textBody?.paragraphs[0]?.align ?? 'left'}`

  if (element.kind === 'shape') {
    return `    <div class="ppt-element ppt-shape ppt-shape-${element.shape}" data-ppt-element="${escapeHtml(element.id)}" style="${[...style, exportShapeStyle(element), textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
  }

  return `    <div class="ppt-element ppt-text" data-ppt-element="${escapeHtml(element.id)}" style="${[...style, textStyle, paragraphStyle].filter(Boolean).join(';')}">${text}</div>`
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
    '.ppt-text{align-items:flex-start;padding:0;}',
    '.ppt-shape{border-radius:24px;}',
    '.ppt-shape-ellipse{border-radius:999px;}',
    '.ppt-shape-diamond{transform:rotate(45deg);}',
    '@media print{body{background:#fff;}.ppt-deck{display:block;padding:0;}.ppt-slide{break-after:page;page-break-after:always;}}',
  ].join('\n')
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
