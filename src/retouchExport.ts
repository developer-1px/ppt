import slideThemeCss from './slideTheme.css?raw'
import { SAMPLE_DECK } from './sampleDeck'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  getRect,
  rectEquals,
  type RetouchDeck,
  type RetouchPatchManifest,
  type SlideBlock,
} from './retouchModel'
import {
  HTML_SLIDE_ATTRIBUTES,
  HTML_SLIDE_CLASSES,
  htmlSlideBlockAttributes,
  htmlSlideRootAttributes,
} from './htmlSlideContract'

export function exportRetouchDeck(deck: RetouchDeck) {
  const patchManifest = buildRetouchPatchManifest(SAMPLE_DECK, deck)
  const css = [
    ":root{--sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-family:var(--sans);color:#111827;background:#f3f4f6;font-synthesis:none;text-rendering:optimizeLegibility;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;}",
    '*{box-sizing:border-box;}',
    'body{margin:0;}',
    '@page{size:16in 9in;margin:0;}',
    '.deck{display:grid;gap:32px;padding:32px;}',
    '.slide{position:relative;container-type:inline-size;width:1280px;height:720px;overflow:hidden;background:#fff;border:1px solid #d1d5db;}',
    '@media print{body{background:#fff;}.deck{display:block;padding:0;}.slide{border:0;break-after:page;page-break-after:always;}}',
    `[${HTML_SLIDE_ATTRIBUTES.block}]{position:absolute;box-sizing:border-box;margin:0;overflow:visible;overflow-wrap:anywhere;white-space:pre-wrap;}`,
    `.${HTML_SLIDE_CLASSES.blockText}{display:block;width:100%;min-width:0;max-width:100%;color:inherit;font:inherit;letter-spacing:inherit;line-height:inherit;overflow-wrap:inherit;text-align:inherit;white-space:inherit;}`,
    slideThemeCss.trim(),
  ].join('\n')

  const slides = deck.slides
    .map((slide) => {
      const baseSlide = SAMPLE_DECK.slides.find(
        (candidate) => candidate.id === slide.id,
      )
      const blocks = slide.blocks
        .map((block, blockIndex) => {
          const baseBlock = baseSlide?.blocks.find(
            (candidate) => candidate.id === block.id,
          )
          const text = escapeHtml(block.text)
          const style = [
            `left:${slidePercent(block.x, SLIDE_WIDTH)}`,
            `top:${slidePercent(block.y, SLIDE_HEIGHT)}`,
            `width:${slidePercent(block.width, SLIDE_WIDTH)}`,
            exportBlockMinimumHeight(block, baseBlock),
          ]
            .filter(Boolean)
            .join(';')

          return `    <${block.tag} ${htmlAttributes(htmlSlideBlockAttributes(block, blockIndex))} class="${escapeHtml(block.className)}" style="${style}"><span class="${HTML_SLIDE_CLASSES.blockText}">${text}</span></${block.tag}>`
        })
        .join('\n')

      return `  <section ${htmlAttributes(htmlSlideRootAttributes(slide.id))} class="slide" style="--accent:${escapeHtml(slide.accent)}">\n${blocks}\n  </section>`
    })
    .join('\n')

  return [
    '<!doctype html>',
    '<html lang="en">',
    '<head>',
    '  <meta charset="utf-8" />',
    '  <meta name="viewport" content="width=device-width, initial-scale=1" />',
    '  <title>Retouched Slides</title>',
    '  <style>',
    indent(css, 4),
    '  </style>',
    '</head>',
    '<body>',
    '<main class="deck">',
    slides,
    '</main>',
    '  <script type="application/json" data-retouch-patch>',
    indent(escapeScriptJson(JSON.stringify(patchManifest, null, 2)), 4),
    '  </script>',
    '</body>',
    '</html>',
    '',
  ].join('\n')
}

function exportBlockMinimumHeight(block: SlideBlock, baseBlock: SlideBlock | undefined) {
  if (block.text.length === 0) {
    return `min-height:${slidePercent(EMPTY_TEXT_BOX_HEIGHT, SLIDE_HEIGHT)}`
  }

  if (
    baseBlock !== undefined &&
    block.text === baseBlock.text &&
    (block.height !== baseBlock.height || block.width !== baseBlock.width)
  ) {
    return `min-height:${slidePercent(block.height, SLIDE_HEIGHT)}`
  }

  return ''
}

function slidePercent(value: number, total: number) {
  return `${Number(((value / total) * 100).toFixed(4))}%`
}

function buildRetouchPatchManifest(
  baseDeck: RetouchDeck,
  deck: RetouchDeck,
): RetouchPatchManifest {
  const text: RetouchPatchManifest['text'] = []
  const layout: RetouchPatchManifest['layout'] = []

  for (const slide of deck.slides) {
    const baseSlide = baseDeck.slides.find((candidate) => candidate.id === slide.id)

    for (const block of slide.blocks) {
      const baseBlock = baseSlide?.blocks.find(
        (candidate) => candidate.id === block.id,
      )
      const rect = getRect(block)
      const baseRect = baseBlock !== undefined ? getRect(baseBlock) : null

      if (baseBlock === undefined || block.text !== baseBlock.text) {
        text.push({
          slideId: slide.id,
          blockId: block.id,
          text: block.text,
        })
      }

      if (baseRect === null || !rectEquals(rect, baseRect)) {
        layout.push({
          slideId: slide.id,
          blockId: block.id,
          rect,
        })
      }
    }
  }

  return {
    version: 1,
    text,
    layout,
  }
}

function escapeHtml(value: string) {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function htmlAttributes(attributes: Record<string, string>) {
  const htmlParts: string[] = []

  for (const name in attributes) {
    htmlParts.push(`${name}="${escapeHtml(attributes[name])}"`)
  }

  return htmlParts.join(' ')
}

function escapeScriptJson(value: string) {
  return value
    .replaceAll('&', '\\u0026')
    .replaceAll('<', '\\u003c')
    .replaceAll('>', '\\u003e')
}

function indent(value: string, spaces: number) {
  const prefix = ' '.repeat(spaces)

  return value
    .split('\n')
    .map((line) => `${prefix}${line}`)
    .join('\n')
}
