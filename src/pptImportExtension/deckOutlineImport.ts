import {
  PPT_DEFAULT_THEME_ID,
  PPT_TITLE_BODY_LAYOUT_ID,
  type PPTParagraph,
  type PPTSlide,
  type PPTTextBody,
} from '../pptModel'
import {
  getPPTCanvasDataTransferText,
} from '../pptCanvasAppAffordanceAdapter'
import {
  getSlideEditMarkdownDeckSource,
  getSlideEditMarkdownDeckSourceFromDataTransfer,
  type SlideEditMarkdownDeckSource,
  type SlideEditMarkdownSlideBlock,
  type SlideEditMarkdownSlideSource,
} from '../pptSlideEditAffordanceAdapter'

export const PPT_DECK_MARKDOWN_OUTLINE_IMPORT_MODEL =
  'ppt-deck-markdown-outline-import'
export const PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT =
  'text-markdown-ppt-outline'

export type PPTDeckMarkdownOutlineSlideSource = {
  body: PPTTextBody
  notes?: string
  title: string
}

export type PPTDeckMarkdownOutlineSource = {
  format: typeof PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT
  slideCount: number
  slides: readonly PPTDeckMarkdownOutlineSlideSource[]
  textLength: number
  title: string
}

type PPTMarkdownOutlineHeading = {
  index: number
  level: number
  text: string
}

const PPT_MARKDOWN_OUTLINE_TITLE_MAX_LENGTH = 80
const PPT_MARKDOWN_OUTLINE_BODY_MAX_PARAGRAPHS = 8
const PPT_MARKDOWN_OUTLINE_BODY_MAX_LENGTH = 140

export function getPPTDeckMarkdownOutlineSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): PPTDeckMarkdownOutlineSource | null {
  if (!dataTransfer) {
    return null
  }

  const slideEditSource = getSlideEditMarkdownDeckSourceFromDataTransfer({
    dataTransfer,
  })

  if (slideEditSource) {
    return createPPTDeckMarkdownOutlineSourceFromSlideEditSource(
      slideEditSource,
    )
  }

  return getPPTDeckMarkdownOutlineSourceFromText(
    getPPTCanvasDataTransferText({ dataTransfer, mimeType: 'text/markdown' }) ||
      getPPTCanvasDataTransferText({
        dataTransfer,
        mimeType: 'text/x-markdown',
      }) ||
      getPPTCanvasDataTransferText({ dataTransfer, mimeType: 'text/plain' }),
  )
}

export function getPPTDeckMarkdownOutlineSourceFromText(
  text: string,
): PPTDeckMarkdownOutlineSource | null {
  const normalized = text.replace(/\r\n?/g, '\n').trim()

  if (!normalized) {
    return null
  }

  const slideEditSource = getSlideEditMarkdownDeckSource({
    markdown: text,
    sourceType: 'text/plain',
  })

  if (slideEditSource) {
    return createPPTDeckMarkdownOutlineSourceFromSlideEditSource(
      slideEditSource,
    )
  }

  const parsed =
    getPPTDeckMarkdownOutlineFromHeadings(normalized) ??
      getPPTDeckMarkdownOutlineFromSeparators(normalized)

  if (!parsed || parsed.slides.length < 2) {
    return null
  }

  return {
    format: PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT,
    slideCount: parsed.slides.length,
    slides: parsed.slides,
    textLength: normalized.length,
    title: parsed.title,
  }
}

export function createPPTDeckMarkdownOutlineSlides({
  createSlideId,
  source,
}: {
  createSlideId: (prefix: string) => string
  source: PPTDeckMarkdownOutlineSource
}): PPTSlide[] {
  return source.slides.map((slideSource, index) => {
    const slideId = createSlideId('slide')
    const titleId = `${slideId}-title`
    const bodyId = `${slideId}-body`
    const bodyParagraphs = slideSource.body.paragraphs
    const elements: PPTSlide['elements'] = [{
      geometry: {
        h: 92,
        w: 1120,
        x: 80,
        y: 64,
      },
      id: titleId,
      kind: 'textBox',
      name: 'Title',
      style: {
        color: '#111827',
        fontSize: 44,
        fontWeight: 'bold',
      },
      textBody: {
        paragraphs: [{
          runs: [{ text: slideSource.title }],
        }],
      },
      textAutoFit: 'resizeShapeToFitText',
    }]

    if (bodyParagraphs.length > 0) {
      elements.push({
        geometry: {
          h: 430,
          w: 1088,
          x: 96,
          y: 184,
        },
        id: bodyId,
        kind: 'textBox',
        name: 'Body',
        style: {
          color: '#1f2937',
          fontSize: 28,
          textInset: {
            bottom: 0,
            left: 0,
            right: 0,
            top: 0,
          },
        },
        textBody: slideSource.body,
      })
    }

    return {
      background: { color: '#ffffff' },
      elements,
      id: slideId,
      layoutId: PPT_TITLE_BODY_LAYOUT_ID,
      name: slideSource.title || `Slide ${index + 1}`,
      ...(slideSource.notes ? { notes: slideSource.notes } : {}),
      themeId: PPT_DEFAULT_THEME_ID,
    }
  })
}

function createPPTDeckMarkdownOutlineSourceFromSlideEditSource(
  source: SlideEditMarkdownDeckSource,
): PPTDeckMarkdownOutlineSource | null {
  const slides = source.slides
    .map(createPPTDeckMarkdownOutlineSlideSourceFromSlideEditSource)
    .filter((slide): slide is PPTDeckMarkdownOutlineSlideSource =>
      slide !== null)

  return slides.length >= 2
    ? {
        format: PPT_DECK_MARKDOWN_OUTLINE_IMPORT_FORMAT,
        slideCount: slides.length,
        slides,
        textLength: source.payloadLength,
        title: source.deckTitle ?? 'Markdown Outline',
      }
    : null
}

function createPPTDeckMarkdownOutlineSlideSourceFromSlideEditSource(
  source: SlideEditMarkdownSlideSource,
): PPTDeckMarkdownOutlineSlideSource | null {
  const paragraphs = createPPTMarkdownOutlineParagraphsFromSlideEditBlocks(
    source.body,
  )
  const notes = readPPTMarkdownOutlineTextFromSlideEditBlocks(source.notes)
  const title = normalizePPTMarkdownOutlineText(source.title)
    .slice(0, PPT_MARKDOWN_OUTLINE_TITLE_MAX_LENGTH)
    .trim()

  if (!title) {
    return null
  }

  return {
    body: {
      paragraphs: paragraphs.slice(0, PPT_MARKDOWN_OUTLINE_BODY_MAX_PARAGRAPHS),
    },
    ...(notes ? { notes } : {}),
    title,
  }
}

function createPPTMarkdownOutlineParagraphsFromSlideEditBlocks(
  blocks: readonly SlideEditMarkdownSlideBlock[],
): PPTParagraph[] {
  return blocks.flatMap((block) => {
    if (block.kind === 'paragraph') {
      return block.markdown
        .split('\n')
        .flatMap((line) => {
          const paragraph = createPPTMarkdownOutlineParagraph(line)

          return paragraph ? [paragraph] : []
        })
    }

    return block.items.flatMap((item) => {
      const text = normalizePPTMarkdownOutlineText(item.text)
        .slice(0, PPT_MARKDOWN_OUTLINE_BODY_MAX_LENGTH)
        .trim()

      return text
        ? [{
            bullet: block.kind === 'ordered-list'
              ? 'numbered' as const
              : 'bullet' as const,
            lineHeight: 1.2,
            runs: [{ text }],
            spacingAfter: 10,
          }]
        : []
    })
  })
}

function readPPTMarkdownOutlineTextFromSlideEditBlocks(
  blocks: readonly SlideEditMarkdownSlideBlock[],
) {
  const lines = blocks.flatMap((block) => {
    if (block.kind === 'paragraph') {
      return [block.text]
    }

    return block.items.map((item) => item.text)
  })
    .map(normalizePPTMarkdownOutlineText)
    .filter(Boolean)

  return lines.join('\n')
}

function getPPTDeckMarkdownOutlineFromHeadings(text: string) {
  const lines = text.split('\n')
  const headings = lines.flatMap((line, index): PPTMarkdownOutlineHeading[] => {
    const match = line.trim().match(/^(#{1,6})\s+(.+?)\s*#*$/)

    return match
      ? [{
          index,
          level: match[1].length,
          text: normalizePPTMarkdownOutlineText(match[2]),
        }]
      : []
  })
  const levelTwoHeadings = headings.filter((heading) => heading.level === 2)
  const slideHeadings = levelTwoHeadings.length >= 2
    ? levelTwoHeadings
    : headings.filter((heading) => heading.level === 1)

  if (slideHeadings.length < 2) {
    return null
  }

  const firstSlideIndex = slideHeadings[0].index
  const deckTitle = headings.find((heading) =>
    heading.level === 1 && heading.index < firstSlideIndex)?.text
  const slides = slideHeadings
    .map((heading, index): PPTDeckMarkdownOutlineSlideSource | null => {
      const nextHeading = slideHeadings[index + 1]
      const bodyLines = lines.slice(
        heading.index + 1,
        nextHeading?.index ?? lines.length,
      )

      return createPPTDeckMarkdownOutlineSlideSource({
        bodyLines,
        title: heading.text,
      })
    })
    .filter((slide): slide is PPTDeckMarkdownOutlineSlideSource =>
      slide !== null)

  return slides.length >= 2
    ? {
        slides,
        title: deckTitle ?? 'Markdown Outline',
      }
    : null
}

function getPPTDeckMarkdownOutlineFromSeparators(text: string) {
  const sections = text
    .split(/\n\s*(?:---|\*\*\*)\s*\n/g)
    .map((section) => section.trim())
    .filter(Boolean)

  if (sections.length < 2) {
    return null
  }

  const slides = sections
    .map((section, index): PPTDeckMarkdownOutlineSlideSource | null => {
      const [rawTitle = '', ...bodyLines] = section.split('\n')
      const title = normalizePPTMarkdownOutlineText(
        rawTitle.replace(/^#{1,6}\s+/, ''),
      ) || `Slide ${index + 1}`

      return createPPTDeckMarkdownOutlineSlideSource({
        bodyLines,
        title,
      })
    })
    .filter((slide): slide is PPTDeckMarkdownOutlineSlideSource =>
      slide !== null)

  return slides.length >= 2
    ? {
        slides,
        title: 'Markdown Outline',
      }
    : null
}

function createPPTDeckMarkdownOutlineSlideSource({
  bodyLines,
  title,
}: {
  bodyLines: readonly string[]
  title: string
}): PPTDeckMarkdownOutlineSlideSource | null {
  const notes: string[] = []
  const paragraphs: PPTParagraph[] = []
  let collectingNotes = false

  for (const rawLine of bodyLines) {
    const line = rawLine.trim()

    if (!line) {
      continue
    }

    const notesMatch = line.match(/^(?:speaker\s+notes?|notes?)\s*:\s*(.*)$/i)

    if (notesMatch) {
      collectingNotes = true
      const firstNote = notesMatch[1].trim()

      if (firstNote) {
        notes.push(normalizePPTMarkdownOutlineText(firstNote))
      }
      continue
    }

    if (collectingNotes) {
      notes.push(normalizePPTMarkdownOutlineText(line))
      continue
    }

    const paragraph = createPPTMarkdownOutlineParagraph(line)

    if (paragraph) {
      paragraphs.push(paragraph)
    }
  }

  const normalizedTitle = normalizePPTMarkdownOutlineText(title)
    .slice(0, PPT_MARKDOWN_OUTLINE_TITLE_MAX_LENGTH)
    .trim()

  if (!normalizedTitle) {
    return null
  }

  return {
    body: {
      paragraphs: paragraphs.slice(0, PPT_MARKDOWN_OUTLINE_BODY_MAX_PARAGRAPHS),
    },
    ...(notes.length > 0 ? { notes: notes.join('\n') } : {}),
    title: normalizedTitle,
  }
}

function createPPTMarkdownOutlineParagraph(line: string): PPTParagraph | null {
  const heading = line.match(/^#{3,6}\s+(.+?)\s*#*$/)
  const bullet = line.match(/^[-*+]\s+(.+)$/)
  const numbered = line.match(/^\d+[.)]\s+(.+)$/)
  const rawText = heading?.[1] ?? bullet?.[1] ?? numbered?.[1] ?? line
  const text = normalizePPTMarkdownOutlineText(rawText)
    .slice(0, PPT_MARKDOWN_OUTLINE_BODY_MAX_LENGTH)
    .trim()

  if (!text) {
    return null
  }

  return {
    ...(bullet ? { bullet: 'bullet' as const } : {}),
    ...(numbered ? { bullet: 'numbered' as const } : {}),
    lineHeight: 1.2,
    runs: [{
      ...(heading ? { bold: true } : {}),
      text,
    }],
    spacingAfter: 10,
  }
}

function normalizePPTMarkdownOutlineText(text: string) {
  return text
    .replace(/!\[([^\]]*)]\([^)]+\)/g, '$1')
    .replace(/\[([^\]]+)]\([^)]+\)/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/_([^_]+)_/g, '$1')
    .replace(/\\([\\`*_[\]()#+.!|-])/g, '$1')
    .replace(/\s+/g, ' ')
    .trim()
}
