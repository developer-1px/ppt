import {
  getPPTDataImageSourceFromDataTransfer,
  getPPTImageFileFromDataTransfer,
  getPPTSVGImageSourceFromDataTransfer,
  type PPTImageImportFormat,
  type PPTImageImportSource,
} from './imageImport'
import {
  getPPTMediaSourceFromDataTransfer,
  type PPTMediaImportSource,
} from './mediaImport'
import {
  getPPTTableColumnCount,
  getPPTTableFileFromDataTransfer,
  getPPTTableSourceFromDataTransfer,
  type PPTTableImportFormat,
  type PPTTableImportSource,
} from './tableImport'
import {
  getPPTRichTextPasteSourceFromDataTransfer,
  getPPTTextPasteSourcesFromDataTransfer,
  type PPTRichTextPasteSource,
} from './textPasteImport'
import type {
  PPTImage,
  PPTTable,
} from '../pptModel'

export const PPT_IMPORT_CANVAS_FALLBACK_ISSUES = {
  richTextClipboard: 'canvas#257',
} as const

export const PPT_IMPORT_EXTENSION = {
  canvasFallbackIssues: [
    PPT_IMPORT_CANVAS_FALLBACK_ISSUES.richTextClipboard,
  ],
  clipboardActionOrder: [
    'image-file',
    'image-source',
    'table-source',
    'media-source',
    'rich-text-source',
    'text-source',
  ],
  dropActionOrder: [
    'image-file',
    'table-file',
    'table-source',
    'media-source',
  ],
  id: 'ppt-import-extension',
  installUnit: 'src/pptImportExtension',
  title: 'PPT import extension',
} as const

export type PPTImportFallbackIssue =
  (typeof PPT_IMPORT_CANVAS_FALLBACK_ISSUES)[keyof typeof PPT_IMPORT_CANVAS_FALLBACK_ISSUES]

export type PPTImageImportEffect = {
  fallbackIssue?: PPTImportFallbackIssue
  format: PPTImageImportFormat
  mimeType?: string
  model: 'canvas-image-import'
  name: string
  naturalHeight?: number
  naturalWidth?: number
}

export type PPTTableImportEffect = {
  columnCount: number
  fallbackIssue?: PPTImportFallbackIssue
  format: PPTTableImportFormat
  model: 'canvas-table-import'
  name: string
  rowCount: number
}

export type PPTClipboardImportAction =
  | {
      file: Blob & { name?: string }
      kind: 'image-file'
    }
  | {
      kind: 'image-source'
      resolveNaturalSize?: boolean
      source: PPTImageImportSource
    }
  | {
      kind: 'table-source'
      source: PPTTableImportSource
    }
  | {
      kind: 'media-source'
      source: PPTMediaImportSource
    }
  | {
      kind: 'rich-text-source'
      source: PPTRichTextPasteSource
    }
  | {
      kind: 'text-source'
      text: string
    }

export type PPTStageDropImportAction =
  | {
      file: Blob & { name?: string }
      kind: 'image-file'
    }
  | {
      fallbackSource: PPTTableImportSource | null
      file: Blob & { name?: string }
      kind: 'table-file'
    }
  | {
      kind: 'table-source'
      source: PPTTableImportSource
    }
  | {
      kind: 'media-source'
      source: PPTMediaImportSource
    }

export function getPPTClipboardImportActions(
  dataTransfer: DataTransfer | null,
): PPTClipboardImportAction[] {
  const actions: PPTClipboardImportAction[] = []

  const file = getPPTImageFileFromDataTransfer(dataTransfer)

  if (file) {
    return [{ file, kind: 'image-file' }]
  }

  const svgImageSource = getPPTSVGImageSourceFromDataTransfer(dataTransfer)

  if (svgImageSource) {
    return [{ kind: 'image-source', source: svgImageSource }]
  }

  const dataImageSource = getPPTDataImageSourceFromDataTransfer(dataTransfer)

  if (dataImageSource) {
    return [{
      kind: 'image-source',
      resolveNaturalSize: true,
      source: dataImageSource,
    }]
  }

  const tableSource = getPPTTableSourceFromDataTransfer(dataTransfer)

  if (tableSource) {
    return [{ kind: 'table-source', source: tableSource }]
  }

  const mediaSource = getPPTMediaSourceFromDataTransfer(dataTransfer)

  if (mediaSource) {
    actions.push({ kind: 'media-source', source: mediaSource })
  }

  const richTextSource = getPPTRichTextPasteSourceFromDataTransfer(dataTransfer)

  if (richTextSource) {
    actions.push({ kind: 'rich-text-source', source: richTextSource })
  }

  actions.push(
    ...getPPTTextPasteSourcesFromDataTransfer(dataTransfer)
      .map((text): PPTClipboardImportAction => ({ kind: 'text-source', text })),
  )

  return actions
}

export function getPPTStageDropImportAction(
  dataTransfer: DataTransfer | null,
): PPTStageDropImportAction | null {
  const imageFile = getPPTImageFileFromDataTransfer(dataTransfer)

  if (imageFile) {
    return { file: imageFile, kind: 'image-file' }
  }

  const tableFile = getPPTTableFileFromDataTransfer(dataTransfer)
  const tableSource = getPPTTableSourceFromDataTransfer(dataTransfer)

  if (tableFile) {
    return {
      fallbackSource: tableSource,
      file: tableFile,
      kind: 'table-file',
    }
  }

  if (tableSource) {
    return { kind: 'table-source', source: tableSource }
  }

  const mediaSource = getPPTMediaSourceFromDataTransfer(dataTransfer)

  return mediaSource ? { kind: 'media-source', source: mediaSource } : null
}

export function canHandlePPTStageDropImport(dataTransfer: DataTransfer | null) {
  return getPPTStageDropImportAction(dataTransfer) !== null
}

export function createPPTImageImportEffect({
  element,
  source,
}: {
  element: PPTImage
  source: PPTImageImportSource
}): PPTImageImportEffect {
  return {
    fallbackIssue: getPPTImageImportFallbackIssue(),
    format: source.format ?? 'file',
    mimeType: source.mimeType,
    model: 'canvas-image-import',
    name: element.name,
    naturalHeight: source.naturalHeight,
    naturalWidth: source.naturalWidth,
  }
}

export function createPPTTableImportEffect({
  element,
  source,
}: {
  element: PPTTable
  source: PPTTableImportSource
}): PPTTableImportEffect {
  return {
    columnCount: getPPTTableColumnCount(element.rows),
    fallbackIssue: getPPTTableImportFallbackIssue(),
    format: source.format ?? 'text-delimited',
    model: 'canvas-table-import',
    name: element.name,
    rowCount: element.rows.length,
  }
}

export function getPPTImageImportFallbackIssue(): PPTImportFallbackIssue | undefined {
  return undefined
}

export function getPPTTableImportFallbackIssue(): PPTImportFallbackIssue | undefined {
  return undefined
}
