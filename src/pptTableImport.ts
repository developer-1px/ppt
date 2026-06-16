import { clamp } from 'canvas/core'
import {
  getCanvasTableCsvFileFromDataTransfer,
  getCanvasTableCsvFileFromList,
  getCanvasTableCsvSourceFromDataTransfer,
  getCanvasTableCsvSourceFromText,
  readCanvasTableCsvFileSource,
  type CanvasTableImportSource,
} from 'canvas/app/table-import'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTTable,
} from './pptModel'

export type PPTTableImportSource = CanvasTableImportSource

export const PPT_DEFAULT_TABLE_ROWS = [
  ['Metric', 'Current', 'Target'],
  ['Draft time', '2h', '30m'],
  ['Review passes', '4', '1'],
]

const PPT_TABLE_TSV_MIME_TYPES = new Set([
  'text/tab-separated-values',
])
const PPT_TABLE_MAX_COLUMNS = 8
const PPT_TABLE_MAX_ROWS = 12
const PPT_TABLE_MAX_CELL_LENGTH = 80
const PPT_TABLE_CELL_WIDTH = 150
const PPT_TABLE_ROW_HEIGHT = 46
const PPT_TABLE_MIN_WIDTH = 260
const PPT_TABLE_MAX_WIDTH = 980
const PPT_TABLE_MIN_HEIGHT = 120
const PPT_TABLE_MAX_HEIGHT = 560

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
  const width = clamp(columnCount * PPT_TABLE_CELL_WIDTH, PPT_TABLE_MIN_WIDTH, PPT_TABLE_MAX_WIDTH)
  const height = clamp(rowCount * PPT_TABLE_ROW_HEIGHT, PPT_TABLE_MIN_HEIGHT, PPT_TABLE_MAX_HEIGHT)

  return {
    geometry: {
      h: height,
      w: width,
      x: clamp(point.x - width / 2, 0, PPT_SLIDE_WIDTH - width),
      y: clamp(point.y - height / 2, 0, PPT_SLIDE_HEIGHT - height),
    },
    id,
    kind: 'table',
    name,
    rows: normalizedRows,
  }
}

export function getPPTTableFileFromList(files: FileList | null) {
  return getCanvasTableCsvFileFromList(files) ??
    Array.from(files ?? []).find(isPPTTableTsvFile) ??
    null
}

export function getPPTTableFileFromDataTransfer(dataTransfer: DataTransfer | null) {
  return getCanvasTableCsvFileFromDataTransfer(dataTransfer) ??
    getPPTTableFileFromList(dataTransfer?.files ?? null)
}

export function getPPTTableSourceFromDataTransfer(dataTransfer: DataTransfer | null) {
  if (!dataTransfer) {
    return null
  }

  const tabSeparatedText = dataTransfer.getData('text/tab-separated-values')

  if (tabSeparatedText) {
    return getPPTTableSourceFromText(tabSeparatedText)
  }

  return toPPTTableImportSource(
    getCanvasTableCsvSourceFromDataTransfer(dataTransfer),
  )
}

export async function readPPTTableFileSource(file: Blob & { name?: string }) {
  const canvasSource = await readCanvasTableCsvFileSource(file)

  if (canvasSource) {
    return toPPTTableImportSource(canvasSource)
  }

  if (!isPPTTableTsvFile(file)) {
    return null
  }

  return getPPTTableSourceFromText(await readPPTBlobAsText(file), {
    name: file.name,
  })
}

export function getPPTTableSourceFromText(
  text: string,
  options: { name?: string } = {},
): PPTTableImportSource | null {
  const canvasSource = getCanvasTableCsvSourceFromText(text, options)

  if (canvasSource) {
    return toPPTTableImportSource(canvasSource)
  }

  const rows = normalizePPTTableRows(parsePPTTableTextRows(text))

  if (!isPPTTableImportRows(rows)) {
    return null
  }

  return {
    ...(options.name === undefined ? {} : { name: getPPTTableImportName(options.name) }),
    rows,
  }
}

export function stringifyPPTTableRows(rows: readonly (readonly string[])[]) {
  return normalizePPTTableRows(rows)
    .map((row) => row.join('\t'))
    .join('\n')
}

export function normalizePPTTableRows(rows: readonly (readonly string[])[]) {
  const cleaned = rows
    .map((row) => row
      .map((cell) => normalizePPTTableCell(cell))
      .slice(0, PPT_TABLE_MAX_COLUMNS))
    .filter((row) => row.some((cell) => cell.length > 0))
    .slice(0, PPT_TABLE_MAX_ROWS)
  const columnCount = getPPTTableColumnCount(cleaned)

  if (cleaned.length === 0 || columnCount === 0) {
    return PPT_DEFAULT_TABLE_ROWS
  }

  return cleaned.map((row) => [
    ...row,
    ...Array.from({ length: columnCount - row.length }, () => ''),
  ])
}

export function getPPTTableColumnCount(rows: readonly (readonly string[])[]) {
  return Math.max(0, ...rows.map((row) => row.length))
}

function toPPTTableImportSource(
  source: CanvasTableImportSource | null,
): PPTTableImportSource | null {
  if (!source) {
    return null
  }

  const rows = normalizePPTTableRows(source.rows)

  if (!isPPTTableImportRows(rows)) {
    return null
  }

  return {
    ...(source.name === undefined ? {} : { name: getPPTTableImportName(source.name) }),
    rows,
  }
}

function isPPTTableTsvFile(file: Blob & { name?: string }) {
  const mimeType = file.type.toLowerCase()
  const name = file.name?.toLowerCase() ?? ''

  return PPT_TABLE_TSV_MIME_TYPES.has(mimeType) || name.endsWith('.tsv')
}

function isPPTTableImportRows(rows: readonly (readonly string[])[]) {
  const nonEmptyRows = rows.filter((row) =>
    row.some((cell) => cell.trim().length > 0),
  )
  const columnCount = getPPTTableColumnCount(nonEmptyRows)

  return nonEmptyRows.length >= 2 && columnCount >= 2
}

function parsePPTTableTextRows(text: string) {
  const delimiter = text.includes('\t') ? '\t' : ','

  return parsePPTDelimitedRows(text, delimiter)
}

function parsePPTDelimitedRows(text: string, delimiter: string) {
  const rows: string[][] = []
  let row: string[] = []
  let cell = ''
  let inQuotes = false

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index]

    if (inQuotes) {
      if (char === '"' && text[index + 1] === '"') {
        cell += '"'
        index += 1
      } else if (char === '"') {
        inQuotes = false
      } else {
        cell += char
      }
      continue
    }

    if (char === '"') {
      inQuotes = true
    } else if (char === delimiter) {
      row.push(cell)
      cell = ''
    } else if (char === '\r' || char === '\n') {
      row.push(cell)
      rows.push(row)
      row = []
      cell = ''

      if (char === '\r' && text[index + 1] === '\n') {
        index += 1
      }
    } else {
      cell += char
    }
  }

  row.push(cell)
  rows.push(row)

  return rows
}

function normalizePPTTableCell(value: string) {
  return value.trim().slice(0, PPT_TABLE_MAX_CELL_LENGTH)
}

function getPPTTableImportName(name: string) {
  return name.replace(/\.[^.]+$/, '') || 'Table'
}

function readPPTBlobAsText(blob: Blob) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()

    reader.addEventListener('load', () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result)
        return
      }

      reject(new Error('Expected table text'))
    })
    reader.addEventListener('error', () => {
      reject(reader.error ?? new Error('Could not read table file'))
    })
    reader.readAsText(blob)
  })
}
