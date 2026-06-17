import {
  createCanvasDataTransferImportActionPlan,
} from 'canvas/app/data-transfer-import-actions'
import { CANVAS_IMAGE_IMPORT_MODEL } from 'canvas/app/image-import'
import { CANVAS_TABLE_IMPORT_MODEL } from 'canvas/app/table-import'
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

export const PPT_IMPORT_EXTENSION = {
  canvasFallbackIssues: [],
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

export type PPTImageImportEffect = {
  format: PPTImageImportFormat
  mimeType?: string
  model: typeof CANVAS_IMAGE_IMPORT_MODEL
  name: string
  naturalHeight?: number
  naturalWidth?: number
}

export type PPTTableImportEffect = {
  columnCount: number
  format: PPTTableImportFormat
  model: typeof CANVAS_TABLE_IMPORT_MODEL
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
  return createCanvasDataTransferImportActionPlan<PPTClipboardImportAction>({
    resolvers: [
      {
        mode: 'exclusive',
        resolve: () => {
          const file = getPPTImageFileFromDataTransfer(dataTransfer)

          return file ? { file, kind: 'image-file' } : null
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
    createCanvasDataTransferImportActionPlan<PPTStageDropImportAction>({
      resolvers: [
        {
          mode: 'exclusive',
          resolve: () => {
            const file = getPPTImageFileFromDataTransfer(dataTransfer)

            return file ? { file, kind: 'image-file' } : null
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
  element,
  source,
}: {
  element: PPTImage
  source: PPTImageImportSource
}): PPTImageImportEffect {
  return {
    format: source.format ?? 'file',
    mimeType: source.mimeType,
    model: CANVAS_IMAGE_IMPORT_MODEL,
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
    format: source.format ?? 'text-delimited',
    model: CANVAS_TABLE_IMPORT_MODEL,
    name: element.name,
    rowCount: element.rows.length,
  }
}
