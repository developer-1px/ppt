import { clampPPTCanvasBoundsToFrame } from '../pptCanvasCoreAdapter'
import {
  getPPTCanvasDataTransferText,
  getPPTCanvasTableColumnCount,
  getPPTCanvasTableComponentSize,
  getPPTCanvasTableFileFromDataTransfer,
  getPPTCanvasTableFilesFromDataTransfer,
  getPPTCanvasTableFileFromList,
  getPPTCanvasTableFilesFromList,
  getPPTCanvasTableSourceFromDataTransfer,
  getPPTCanvasTableSourceFromHTML,
  getPPTCanvasTableSourceFromText,
  normalizePPTCanvasTableRows,
  PPT_CANVAS_TABLE_IMPORT_MODEL,
  readPPTCanvasTableFileSource,
  readPPTCanvasTableFileSources,
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
  | 'text-markdown'
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
const PPT_TABLE_BATCH_GAP = 24
const PPT_TABLE_BATCH_MAX_COLUMNS = 2
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

export function createPPTTableElements({
  center,
  createId,
  sources,
}: {
  center: { x: number; y: number }
  createId: (prefix: string) => string
  sources: readonly PPTTableImportSource[]
}) {
  if (sources.length <= 1) {
    return sources.map((source) =>
      createPPTTableElement({
        id: createId('table'),
        name: source.name ?? 'Table',
        point: center,
        rows: source.rows,
      })
    )
  }

  const columns = Math.min(
    PPT_TABLE_BATCH_MAX_COLUMNS,
    Math.ceil(Math.sqrt(sources.length)),
  )
  const rows = Math.ceil(sources.length / columns)
  const sizes = sources.map(getPPTTableImportSourceSize)
  const columnWidths = Array.from({ length: columns }, (_, column) =>
    Math.max(
      ...sizes
        .filter((_, index) => index % columns === column)
        .map((size) => size.w),
    )
  )
  const rowHeights = Array.from({ length: rows }, (_, row) =>
    Math.max(
      ...sizes
        .filter((_, index) => Math.floor(index / columns) === row)
        .map((size) => size.h),
    )
  )
  const totalWidth = columnWidths.reduce((sum, width) => sum + width, 0) +
    PPT_TABLE_BATCH_GAP * Math.max(0, columns - 1)
  const totalHeight = rowHeights.reduce((sum, height) => sum + height, 0) +
    PPT_TABLE_BATCH_GAP * Math.max(0, rows - 1)
  const origin = {
    x: center.x - totalWidth / 2,
    y: center.y - totalHeight / 2,
  }

  return sources.map((source, index) => {
    const column = index % columns
    const row = Math.floor(index / columns)
    const x = origin.x +
      columnWidths.slice(0, column).reduce((sum, width) => sum + width, 0) +
      PPT_TABLE_BATCH_GAP * column +
      columnWidths[column] / 2
    const y = origin.y +
      rowHeights.slice(0, row).reduce((sum, height) => sum + height, 0) +
      PPT_TABLE_BATCH_GAP * row +
      rowHeights[row] / 2

    return createPPTTableElement({
      id: createId('table'),
      name: source.name ?? 'Table',
      point: { x, y },
      rows: source.rows,
    })
  })
}

export function getPPTTableFileFromList(files: FileList | null) {
  return getPPTCanvasTableFileFromList(files)
}

export function getPPTTableFilesFromList(files: FileList | null) {
  return getPPTCanvasTableFilesFromList(files)
}

export function getPPTTableFileFromDataTransfer(dataTransfer: DataTransfer | null) {
  return getPPTCanvasTableFileFromDataTransfer(dataTransfer)
}

export function getPPTTableFilesFromDataTransfer(dataTransfer: DataTransfer | null) {
  return getPPTCanvasTableFilesFromDataTransfer(dataTransfer)
}

export function getPPTTableSourceFromDataTransfer(dataTransfer: DataTransfer | null) {
  const source = getPPTCanvasTableSourceFromDataTransfer(dataTransfer)

  if (source) {
    return createPPTTableImportSourceFromCanvas(source)
  }

  if (!dataTransfer) {
    return null
  }

  return getPPTMarkdownTableSourceFromText(
    getPPTCanvasDataTransferText({ dataTransfer, mimeType: 'text/markdown' }) ||
      getPPTCanvasDataTransferText({ dataTransfer, mimeType: 'text/plain' }),
  )
}

export async function readPPTTableFileSource(file: Blob & { name?: string }) {
  const source = await readPPTCanvasTableFileSource(file)

  if (!source) {
    return null
  }

  return createPPTTableImportSourceFromCanvas(source)
}

export async function readPPTTableFileSources(
  files: readonly (Blob & { name?: string })[],
) {
  return (await readPPTCanvasTableFileSources(files)).map((source) =>
    createPPTTableImportSourceFromCanvas(source)
  )
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

export function getPPTMarkdownTableSourceFromText(
  text: string,
): PPTTableImportSource | null {
  const rows = getPPTMarkdownTableRows(text)

  return isPPTTableImportRows(rows)
    ? {
        format: 'text-markdown',
        name: 'Markdown Table',
        rows,
      }
    : null
}

export function stringifyPPTTableRows(rows: readonly (readonly string[])[]) {
  return normalizePPTTableRows(rows)
    .map((row) => row.join('\t'))
    .join('\n')
}

export function stringifyPPTTableRowsCSV(
  rows: readonly (readonly string[])[],
) {
  return normalizePPTTableRows(rows)
    .map((row) => row.map(formatPPTTableCSVCell).join(','))
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

function formatPPTTableCSVCell(value: string) {
  return /[",\r\n]/.test(value)
    ? `"${value.replace(/"/g, '""')}"`
    : value
}

function getPPTTableImportSourceSize(source: PPTTableImportSource) {
  const normalizedRows = normalizePPTTableRows(source.rows)

  return getPPTCanvasTableComponentSize({
    columnCount: getPPTTableColumnCount(normalizedRows),
    rowCount: normalizedRows.length,
  }, PPT_TABLE_SIZE_OPTIONS)
}

function getPPTMarkdownTableRows(text: string) {
  const lines = text.split(/\r?\n/).map((line) => line.trim())

  for (let index = 0; index < lines.length - 1; index += 1) {
    if (!lines[index] || !lines[index + 1]) {
      continue
    }

    const header = parsePPTMarkdownTableRow(lines[index])
    const separator = parsePPTMarkdownTableRow(lines[index + 1])

    if (
      header.length < 2 ||
      separator.length !== header.length ||
      !separator.every(isPPTMarkdownTableSeparatorCell)
    ) {
      continue
    }

    const bodyRows: string[][] = []

    for (let rowIndex = index + 2; rowIndex < lines.length; rowIndex += 1) {
      if (!lines[rowIndex]) {
        break
      }

      const row = parsePPTMarkdownTableRow(lines[rowIndex])

      if (row.length !== header.length) {
        break
      }

      bodyRows.push(row)
    }

    if (bodyRows.length === 0) {
      continue
    }

    return [header, ...bodyRows]
  }

  return []
}

function parsePPTMarkdownTableRow(line: string) {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '')
  const cells: string[] = []
  let cell = ''
  let escaped = false

  for (const char of trimmed) {
    if (escaped) {
      cell += char
      escaped = false
      continue
    }

    if (char === '\\') {
      escaped = true
      continue
    }

    if (char === '|') {
      cells.push(normalizePPTMarkdownTableCell(cell))
      cell = ''
      continue
    }

    cell += char
  }

  cells.push(normalizePPTMarkdownTableCell(cell))

  return cells
}

function normalizePPTMarkdownTableCell(value: string) {
  return value
    .trim()
    .replace(/\[([^\]]+)]\([^)]+\)/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
}

function isPPTMarkdownTableSeparatorCell(value: string) {
  return /^:?-{3,}:?$/.test(value.replace(/\s+/g, ''))
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
    format === 'text-markdown' ||
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
    format === 'text-markdown' ||
    format === 'text-tsv'
  ) {
    return format
  }

  return undefined
}
