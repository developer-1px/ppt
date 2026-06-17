import { clampPPTCanvasBoundsToFrame } from '../pptCanvasCoreAdapter'
import {
  getPPTCanvasTableColumnCount,
  getPPTCanvasTableComponentSize,
  getPPTCanvasTableFileFromDataTransfer,
  getPPTCanvasTableFileFromList,
  getPPTCanvasTableSourceFromDataTransfer,
  getPPTCanvasTableSourceFromHTML,
  getPPTCanvasTableSourceFromText,
  normalizePPTCanvasTableRows,
  PPT_CANVAS_TABLE_IMPORT_MODEL,
  readPPTCanvasTableFileSource,
  type PPTCanvasTableImportFormat,
  type PPTCanvasTableImportSource,
} from '../pptCanvasAppAffordanceAdapter'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTTable,
} from '../pptModel'

export type PPTTableImportFormat =
  | 'canvas-csv'
  | 'default'
  | 'text-delimited'
  | 'text-html'
  | 'text-tsv'
export type PPTTableImportSource = {
  format?: PPTTableImportFormat
  name?: string
  rows: readonly (readonly string[])[]
}
export const PPT_TABLE_IMPORT_MODEL = PPT_CANVAS_TABLE_IMPORT_MODEL

export const PPT_DEFAULT_TABLE_ROWS = [
  ['Metric', 'Current', 'Target'],
  ['Draft time', '2h', '30m'],
  ['Review passes', '4', '1'],
]

const PPT_TABLE_MAX_COLUMNS = 8
const PPT_TABLE_MAX_ROWS = 12
const PPT_TABLE_MAX_CELL_LENGTH = 80
const PPT_TABLE_SIZE_OPTIONS = {
  cellSize: { h: 46, w: 150 },
  maxSize: { h: 560, w: 980 },
  minSize: { h: 120, w: 260 },
} as const

export function createPPTTableElement({
  id,
  name = 'Table',
  point,
  rows,
}: {
  id: string
  name?: string
  point: { x: number; y: number }
  rows: readonly (readonly string[])[]
}): PPTTable {
  const normalizedRows = normalizePPTTableRows(rows)
  const columnCount = getPPTTableColumnCount(normalizedRows)
  const rowCount = normalizedRows.length
  const size = getPPTCanvasTableComponentSize({
    columnCount,
    rowCount,
  }, PPT_TABLE_SIZE_OPTIONS)
  const geometry = clampPPTCanvasBoundsToFrame({
    bounds: {
      h: size.h,
      w: size.w,
      x: point.x - size.w / 2,
      y: point.y - size.h / 2,
    },
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
  })

  return {
    geometry,
    id,
    kind: 'table',
    name,
    rows: normalizedRows,
  }
}

export function getPPTTableFileFromList(files: FileList | null) {
  return getPPTCanvasTableFileFromList(files)
}

export function getPPTTableFileFromDataTransfer(dataTransfer: DataTransfer | null) {
  return getPPTCanvasTableFileFromDataTransfer(dataTransfer)
}

export function getPPTTableSourceFromDataTransfer(dataTransfer: DataTransfer | null) {
  const source = getPPTCanvasTableSourceFromDataTransfer(dataTransfer)

  if (!source) {
    return null
  }

  return createPPTTableImportSourceFromCanvas(source)
}

export async function readPPTTableFileSource(file: Blob & { name?: string }) {
  const source = await readPPTCanvasTableFileSource(file)

  if (!source) {
    return null
  }

  return createPPTTableImportSourceFromCanvas(source)
}

export function getPPTTableSourceFromText(
  text: string,
  options: { format?: PPTTableImportFormat; name?: string } = {},
): PPTTableImportSource | null {
  const source = getPPTCanvasTableSourceFromText(text, {
    format: getPPTCanvasTableImportFormat(options.format),
    name: options.name,
  })

  if (!source) {
    return null
  }

  return createPPTTableImportSourceFromCanvas(source, {
    format: options.format,
  })
}

export function getPPTTableSourceFromHTML(value: string) {
  const source = getPPTCanvasTableSourceFromHTML(value)

  if (!source) {
    return null
  }

  return createPPTTableImportSourceFromCanvas(source)
}

export function stringifyPPTTableRows(rows: readonly (readonly string[])[]) {
  return normalizePPTTableRows(rows)
    .map((row) => row.join('\t'))
    .join('\n')
}

export function normalizePPTTableRows(rows: readonly (readonly string[])[]) {
  return normalizePPTCanvasTableRows(rows, {
    fallbackRows: PPT_DEFAULT_TABLE_ROWS,
    maxCellLength: PPT_TABLE_MAX_CELL_LENGTH,
    maxColumns: PPT_TABLE_MAX_COLUMNS,
    maxRows: PPT_TABLE_MAX_ROWS,
  })
}

export function getPPTTableColumnCount(rows: readonly (readonly string[])[]) {
  return getPPTCanvasTableColumnCount(rows)
}

function isPPTTableImportRows(rows: readonly (readonly string[])[]) {
  const nonEmptyRows = rows.filter((row) =>
    row.some((cell) => cell.trim().length > 0),
  )
  const columnCount = getPPTTableColumnCount(nonEmptyRows)

  return nonEmptyRows.length >= 2 && columnCount >= 2
}

function createPPTTableImportSource(
  rows: readonly (readonly string[])[],
  options: { format?: PPTTableImportFormat; name?: string } = {},
): PPTTableImportSource | null {
  if (!isPPTTableImportRows(rows)) {
    return null
  }

  return {
    ...(options.format === undefined ? {} : { format: options.format }),
    ...(options.name === undefined ? {} : { name: getPPTTableImportName(options.name) }),
    rows: normalizePPTTableRows(rows),
  }
}

function getPPTTableImportName(name: string) {
  return name.replace(/\.[^.]+$/, '') || 'Table'
}

function createPPTTableImportSourceFromCanvas(
  source: PPTCanvasTableImportSource,
  options: { format?: PPTTableImportFormat } = {},
) {
  return createPPTTableImportSource(source.rows, {
    format: options.format ?? getPPTTableImportFormat(source.format),
    name: source.name,
  })
}

function getPPTCanvasTableImportFormat(
  format?: PPTTableImportFormat,
): PPTCanvasTableImportFormat | undefined {
  if (format === 'canvas-csv') {
    return 'text-csv'
  }

  if (
    format === 'text-delimited' ||
    format === 'text-html' ||
    format === 'text-tsv'
  ) {
    return format
  }

  return undefined
}

function getPPTTableImportFormat(
  format?: PPTCanvasTableImportFormat,
): PPTTableImportFormat | undefined {
  if (format === 'text-csv') {
    return 'canvas-csv'
  }

  if (
    format === 'text-delimited' ||
    format === 'text-html' ||
    format === 'text-tsv'
  ) {
    return format
  }

  return undefined
}
