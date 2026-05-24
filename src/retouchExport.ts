import slideThemeCss from './slideTheme.css?raw'
import { SAMPLE_DECK } from './sampleDeck'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  getRect,
  rectEquals,
  type RetouchDeck,
  type RetouchPatchManifest,
  type SlideBlock,
} from './retouchModel'

export function exportRetouchDeck(deck: RetouchDeck) {
  const patchManifest = buildRetouchPatchManifest(SAMPLE_DECK, deck)
  const css = [
    ':root{--sans:Inter,ui-sans-serif,system-ui,sans-serif;font-family:var(--sans);color:#111827;background:#f3f4f6;}',
    'body{margin:0;}',
    '@page{size:16in 9in;margin:0;}',
    '.deck{display:grid;gap:32px;padding:32px;}',
    '.slide{position:relative;container-type:inline-size;width:1280px;height:720px;overflow:hidden;background:#fff;border:1px solid #d1d5db;}',
    '@media print{body{background:#fff;}.deck{display:block;padding:0;}.slide{border:0;break-after:page;page-break-after:always;}}',
    '[data-block]{position:absolute;box-sizing:border-box;margin:0;overflow:visible;overflow-wrap:anywhere;white-space:pre-wrap;}',
    slideThemeCss.trim(),
  ].join('\n')

  const slides = deck.slides
    .map((slide) => {
      const baseSlide = SAMPLE_DECK.slides.find(
        (candidate) => candidate.id === slide.id,
      )
      const blocks = slide.blocks
        .map((block) => {
          const baseBlock = baseSlide?.blocks.find(
            (candidate) => candidate.id === block.id,
          )
          const text = escapeHtml(block.text)
          const style = [
            `left:${block.x}px`,
            `top:${block.y}px`,
            `width:${block.width}px`,
            exportBlockMinimumHeight(block, baseBlock),
          ]
            .filter(Boolean)
            .join(';')

          return `    <${block.tag} data-block="${block.id}" data-role="${block.role}" class="${block.className}" style="${style}">${text}</${block.tag}>`
        })
        .join('\n')

      return `  <section data-slide="${slide.id}" class="slide" style="--accent:${slide.accent}">\n${blocks}\n  </section>`
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
    return `min-height:${EMPTY_TEXT_BOX_HEIGHT}px`
  }

  if (baseBlock && block.text === baseBlock.text && block.height !== baseBlock.height) {
    return `min-height:${block.height}px`
  }

  return ''
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

      if (!baseBlock || block.text !== baseBlock.text) {
        text.push({
          slideId: slide.id,
          blockId: block.id,
          text: block.text,
        })
      }

      if (!baseBlock || !rectEquals(getRect(block), getRect(baseBlock))) {
        layout.push({
          slideId: slide.id,
          blockId: block.id,
          rect: getRect(block),
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
