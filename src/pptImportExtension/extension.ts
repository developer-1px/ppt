import {
  createPPTCanvasDataTransferImportActionPlanFromRegistry,
  createPPTCanvasDataTransferImportRegistry,
  getPPTCanvasDataTransferImportRegistryMetadata,
  type PPTCanvasDataTransferImportRegistryResolver,
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
  getPPTTableFilesFromDataTransfer,
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

const PPT_DATA_TRANSFER_IMPORT_REGISTRY_MODEL =
  'canvas-data-transfer-import-registry'
const PPT_CLIPBOARD_IMPORT_SCOPE = 'clipboard'
const PPT_STAGE_DROP_IMPORT_SCOPE = 'stage-drop'

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
  count: number
  format: PPTTableImportFormat
  model: typeof PPT_TABLE_IMPORT_MODEL
  name: string
  names: string
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
      files: readonly (Blob & { name?: string })[]
      kind: 'table-file-batch'
    }
  | {
      file: Blob & { name?: string }
      kind: 'table-file'
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
      files: readonly (Blob & { name?: string })[]
      kind: 'table-file-batch'
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

type PPTDataTransferImportScope =
  | typeof PPT_CLIPBOARD_IMPORT_SCOPE
  | typeof PPT_STAGE_DROP_IMPORT_SCOPE
type PPTDataTransferImportAction =
  | PPTClipboardImportAction
  | PPTStageDropImportAction

const PPT_DATA_TRANSFER_IMPORT_RESOLVERS = [
  {
    id: 'image-file-batch',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const files = getPPTImageFilesFromDataTransfer(dataTransfer)

      return files.length > 1
        ? { files, kind: 'image-file-batch' }
        : null
    },
    scope: [PPT_CLIPBOARD_IMPORT_SCOPE, PPT_STAGE_DROP_IMPORT_SCOPE],
    supportedFormats: ['Files', 'image/*'],
    title: 'Image file batch',
  },
  {
    id: 'image-file',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const files = getPPTImageFilesFromDataTransfer(dataTransfer)
      const file = files.length === 1
        ? files[0]
        : getPPTImageFileFromDataTransfer(dataTransfer)

      return file ? { file, kind: 'image-file' } : null
    },
    scope: [PPT_CLIPBOARD_IMPORT_SCOPE, PPT_STAGE_DROP_IMPORT_SCOPE],
    supportedFormats: ['Files', 'image/*'],
    title: 'Image file',
  },
  {
    id: 'table-file-batch',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const files = getPPTTableFilesFromDataTransfer(dataTransfer)

      return files.length > 1
        ? { files, kind: 'table-file-batch' }
        : null
    },
    scope: [PPT_CLIPBOARD_IMPORT_SCOPE, PPT_STAGE_DROP_IMPORT_SCOPE],
    supportedFormats: ['Files', 'text/csv', 'text/tab-separated-values'],
    title: 'Table file batch',
  },
  {
    id: 'table-file',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const files = getPPTTableFilesFromDataTransfer(dataTransfer)
      const file = files.length === 1
        ? files[0]
        : getPPTTableFileFromDataTransfer(dataTransfer)

      return file ? { file, kind: 'table-file' } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['Files', 'text/csv', 'text/tab-separated-values'],
    title: 'Table file',
  },
  {
    id: 'table-file',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const file = getPPTTableFileFromDataTransfer(dataTransfer)

      return file
        ? {
            fallbackSource: getPPTTableSourceFromDataTransfer(dataTransfer),
            file,
            kind: 'table-file',
          }
        : null
    },
    scope: PPT_STAGE_DROP_IMPORT_SCOPE,
    supportedFormats: ['Files', 'text/csv', 'text/tab-separated-values'],
    title: 'Table file',
  },
  {
    id: 'fallback-html-selection-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source =
        getPPTFallbackHTMLSelectionSourceFromDataTransfer(dataTransfer)

      return source
        ? { kind: 'fallback-html-selection-source', source }
        : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Fallback HTML selection',
  },
  {
    id: 'fallback-html-image-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTFallbackHTMLImageSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'fallback-html-image-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Fallback HTML image',
  },
  {
    id: 'fallback-html-shape-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTFallbackHTMLShapeSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'fallback-html-shape-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Fallback HTML shape',
  },
  {
    id: 'fallback-html-table-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTFallbackHTMLTableSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'fallback-html-table-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Fallback HTML table',
  },
  {
    id: 'fallback-html-text-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTFallbackHTMLTextSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'fallback-html-text-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Fallback HTML text',
  },
  {
    id: 'image-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTSVGImageSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'image-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['image/svg+xml', 'text/html'],
    title: 'SVG image source',
  },
  {
    id: 'image-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTDataImageSourceFromDataTransfer(dataTransfer)

      return source
        ? { kind: 'image-source', resolveNaturalSize: true, source }
        : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html'],
    title: 'Data image source',
  },
  {
    id: 'table-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTTableSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'table-source', source } : null
    },
    scope: [PPT_CLIPBOARD_IMPORT_SCOPE, PPT_STAGE_DROP_IMPORT_SCOPE],
    supportedFormats: ['text/html', 'text/markdown', 'text/plain', 'text/csv'],
    title: 'Table source',
  },
  {
    id: 'rich-text-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTRichTextPasteSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'rich-text-source', source } : null
    },
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/html', 'text/markdown', 'text/plain'],
    title: 'Rich text source',
  },
  {
    id: 'media-source',
    mode: 'exclusive',
    resolve: ({ dataTransfer }) => {
      const source = getPPTMediaSourceFromDataTransfer(dataTransfer)

      return source ? { kind: 'media-source', source } : null
    },
    scope: [PPT_CLIPBOARD_IMPORT_SCOPE, PPT_STAGE_DROP_IMPORT_SCOPE],
    supportedFormats: ['text/plain', 'text/uri-list'],
    title: 'Media source',
  },
  {
    id: 'text-source',
    mode: 'append',
    resolve: ({ dataTransfer }) =>
      getPPTTextPasteSourcesFromDataTransfer(dataTransfer)
        .map((text): PPTClipboardImportAction => ({ kind: 'text-source', text })),
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
    supportedFormats: ['text/plain'],
    title: 'Text source',
  },
] satisfies readonly PPTCanvasDataTransferImportRegistryResolver<
  PPTDataTransferImportAction,
  PPTDataTransferImportScope
>[]

const PPT_DATA_TRANSFER_IMPORT_REGISTRY =
  createPPTCanvasDataTransferImportRegistry<
    PPTDataTransferImportAction,
    PPTDataTransferImportScope
  >({
    resolvers: PPT_DATA_TRANSFER_IMPORT_RESOLVERS,
  })

function getPPTDataTransferImportActionOrder(
  scope: PPTDataTransferImportScope,
) {
  return [
    ...new Set(getPPTCanvasDataTransferImportRegistryMetadata({
      registry: PPT_DATA_TRANSFER_IMPORT_REGISTRY,
      scope,
    }).map((metadata) => metadata.id)),
  ]
}

export const PPT_IMPORT_EXTENSION = {
  actionPlanner: PPT_DATA_TRANSFER_IMPORT_REGISTRY_MODEL,
  canvasFallbackIssues: [],
  clipboardActionOrder: getPPTDataTransferImportActionOrder(
    PPT_CLIPBOARD_IMPORT_SCOPE,
  ),
  dropActionOrder: getPPTDataTransferImportActionOrder(
    PPT_STAGE_DROP_IMPORT_SCOPE,
  ),
  id: 'ppt-import-extension',
  installUnit: 'src/pptImportExtension',
  title: 'PPT import extension',
} as const

export function getPPTClipboardImportActions(
  dataTransfer: DataTransfer | null,
): PPTClipboardImportAction[] {
  return createPPTCanvasDataTransferImportActionPlanFromRegistry({
    dataTransfer,
    registry: PPT_DATA_TRANSFER_IMPORT_REGISTRY,
    scope: PPT_CLIPBOARD_IMPORT_SCOPE,
  }) as PPTClipboardImportAction[]
}

export function getPPTStageDropImportAction(
  dataTransfer: DataTransfer | null,
): PPTStageDropImportAction | null {
  const [action = null] = createPPTCanvasDataTransferImportActionPlanFromRegistry({
    dataTransfer,
    registry: PPT_DATA_TRANSFER_IMPORT_REGISTRY,
    scope: PPT_STAGE_DROP_IMPORT_SCOPE,
  }) as PPTStageDropImportAction[]

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
  batch,
  element,
  source,
}: {
  batch?: {
    count: number
    names: readonly string[]
  }
  element: PPTTable
  source: PPTTableImportSource
}): PPTTableImportEffect {
  return {
    columnCount: getPPTTableColumnCount(element.rows),
    count: batch?.count ?? 1,
    format: source.format ?? 'text-delimited',
    model: PPT_TABLE_IMPORT_MODEL,
    name: element.name,
    names: (batch?.names ?? [element.name]).join(', '),
    rowCount: element.rows.length,
  }
}

export { createPPTFallbackHTMLImportEffect }
