import {
  createPPTCanvasDataTransferImportActionPlan,
} from '../pptCanvasAppAffordanceAdapter'
import {
  getPPTDataImageSourceFromDataTransfer,
  getPPTImageFileFromDataTransfer,
  getPPTImageFilesFromDataTransfer,
  getPPTSVGImageSourceFromDataTransfer,
  PPT_IMAGE_IMPORT_MODEL,
  type PPTImageImportFormat,
  type PPTImageImportSource,
} from './imageImport'
import {
  getPPTMediaSourceFromDataTransfer,
  type PPTMediaImportSource,
} from './mediaImport'
import {
  createPPTFallbackHTMLImportEffect,
  getPPTFallbackHTMLImageSourceFromDataTransfer,
  getPPTFallbackHTMLSelectionSourceFromDataTransfer,
  getPPTFallbackHTMLShapeSourceFromDataTransfer,
  getPPTFallbackHTMLTableSourceFromDataTransfer,
  getPPTFallbackHTMLTextSourceFromDataTransfer,
  PPT_FALLBACK_HTML_IMPORT_MODEL,
  type PPTFallbackHTMLImportEffect,
  type PPTFallbackHTMLImageSource,
  type PPTFallbackHTMLSelectionSource,
  type PPTFallbackHTMLShapeSource,
  type PPTFallbackHTMLTableSource,
  type PPTFallbackHTMLTextSource,
} from './pptFallbackHTMLImport'
import {
  getPPTTableColumnCount,
  getPPTTableFileFromDataTransfer,
  getPPTTableSourceFromDataTransfer,
  PPT_TABLE_IMPORT_MODEL,
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

export const PPT_IMPORT_EXTENSION = {
  canvasFallbackIssues: [],
  clipboardActionOrder: [
    'image-file-batch',
    'image-file',
    'fallback-html-selection-source',
    'fallback-html-image-source',
    'fallback-html-shape-source',
    'fallback-html-table-source',
    'fallback-html-text-source',
    'image-source',
    'table-source',
    'media-source',
    'rich-text-source',
    'text-source',
  ],
  dropActionOrder: [
    'image-file-batch',
    'image-file',
    'table-file',
    'table-source',
    'media-source',
  ],
  id: 'ppt-import-extension',
  installUnit: 'src/pptImportExtension',
  title: 'PPT import extension',
} as const

export type PPTImageImportEffect = {
  count: number
  format: PPTImageImportFormat
  mimeType?: string
  model: typeof PPT_IMAGE_IMPORT_MODEL
  name: string
  names: string
  naturalHeight?: number
  naturalWidth?: number
}

export type PPTTableImportEffect = {
  columnCount: number
  format: PPTTableImportFormat
  model: typeof PPT_TABLE_IMPORT_MODEL
  name: string
  rowCount: number
}
export type { PPTFallbackHTMLImportEffect }
export { PPT_FALLBACK_HTML_IMPORT_MODEL }

export type PPTClipboardImportAction =
  | {
      files: readonly (Blob & { name?: string })[]
      kind: 'image-file-batch'
    }
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
      kind: 'fallback-html-selection-source'
      source: PPTFallbackHTMLSelectionSource
    }
  | {
      kind: 'fallback-html-image-source'
      source: PPTFallbackHTMLImageSource
    }
  | {
      kind: 'fallback-html-shape-source'
      source: PPTFallbackHTMLShapeSource
    }
  | {
      kind: 'fallback-html-table-source'
      source: PPTFallbackHTMLTableSource
    }
  | {
      kind: 'fallback-html-text-source'
      source: PPTFallbackHTMLTextSource
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
      files: readonly (Blob & { name?: string })[]
      kind: 'image-file-batch'
    }
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
  return createPPTCanvasDataTransferImportActionPlan<PPTClipboardImportAction>({
    resolvers: [
      {
        mode: 'exclusive',
        resolve: () => {
          const files = getPPTImageFilesFromDataTransfer(dataTransfer)
          const file = files.length === 1
            ? files[0]
            : getPPTImageFileFromDataTransfer(dataTransfer)

          return files.length > 1
            ? { files, kind: 'image-file-batch' }
            : file
              ? { file, kind: 'image-file' }
              : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source =
            getPPTFallbackHTMLSelectionSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'fallback-html-selection-source', source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source =
            getPPTFallbackHTMLImageSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'fallback-html-image-source', source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source =
            getPPTFallbackHTMLShapeSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'fallback-html-shape-source', source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source =
            getPPTFallbackHTMLTableSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'fallback-html-table-source', source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source =
            getPPTFallbackHTMLTextSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'fallback-html-text-source', source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source = getPPTSVGImageSourceFromDataTransfer(dataTransfer)

          return source ? { kind: 'image-source', source } : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source = getPPTDataImageSourceFromDataTransfer(dataTransfer)

          return source
            ? { kind: 'image-source', resolveNaturalSize: true, source }
            : null
        },
      },
      {
        mode: 'exclusive',
        resolve: () => {
          const source = getPPTTableSourceFromDataTransfer(dataTransfer)

          return source ? { kind: 'table-source', source } : null
        },
      },
      {
        mode: 'append',
        resolve: () => {
          const source = getPPTMediaSourceFromDataTransfer(dataTransfer)

          return source ? { kind: 'media-source', source } : null
        },
      },
      {
        mode: 'append',
        resolve: () => {
          const source = getPPTRichTextPasteSourceFromDataTransfer(dataTransfer)

          return source ? { kind: 'rich-text-source', source } : null
        },
      },
      {
        mode: 'append',
        resolve: () => getPPTTextPasteSourcesFromDataTransfer(dataTransfer)
          .map((text): PPTClipboardImportAction => ({ kind: 'text-source', text })),
      },
    ],
  })
}

export function getPPTStageDropImportAction(
  dataTransfer: DataTransfer | null,
): PPTStageDropImportAction | null {
  const [action = null] =
    createPPTCanvasDataTransferImportActionPlan<PPTStageDropImportAction>({
      resolvers: [
        {
          mode: 'exclusive',
          resolve: () => {
            const files = getPPTImageFilesFromDataTransfer(dataTransfer)
            const file = files.length === 1
              ? files[0]
              : getPPTImageFileFromDataTransfer(dataTransfer)

            return files.length > 1
              ? { files, kind: 'image-file-batch' }
              : file
                ? { file, kind: 'image-file' }
                : null
          },
        },
        {
          mode: 'exclusive',
          resolve: () => {
            const file = getPPTTableFileFromDataTransfer(dataTransfer)

            return file
              ? {
                  fallbackSource:
                    getPPTTableSourceFromDataTransfer(dataTransfer),
                  file,
                  kind: 'table-file',
                }
              : null
          },
        },
        {
          mode: 'exclusive',
          resolve: () => {
            const source = getPPTTableSourceFromDataTransfer(dataTransfer)

            return source ? { kind: 'table-source', source } : null
          },
        },
        {
          mode: 'exclusive',
          resolve: () => {
            const source = getPPTMediaSourceFromDataTransfer(dataTransfer)

            return source ? { kind: 'media-source', source } : null
          },
        },
      ],
    })

  return action
}

export function canHandlePPTStageDropImport(dataTransfer: DataTransfer | null) {
  return getPPTStageDropImportAction(dataTransfer) !== null
}

export function createPPTImageImportEffect({
  batch,
  element,
  source,
}: {
  batch?: {
    count: number
    names: readonly string[]
  }
  element: PPTImage
  source: PPTImageImportSource
}): PPTImageImportEffect {
  return {
    count: batch?.count ?? 1,
    format: source.format ?? 'file',
    mimeType: source.mimeType,
    model: PPT_IMAGE_IMPORT_MODEL,
    name: element.name,
    names: (batch?.names ?? [element.name]).join(', '),
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
    format: source.format ?? 'text-delimited',
    model: PPT_TABLE_IMPORT_MODEL,
    name: element.name,
    rowCount: element.rows.length,
  }
}

export { createPPTFallbackHTMLImportEffect }
