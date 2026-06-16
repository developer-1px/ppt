import { clamp } from 'canvas/core'
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

export const PPT_DEFAULT_TABLE_ROWS = [
  ['Metric', 'Current', 'Target'],
  ['Draft time', '2h', '30m'],
  ['Review passes', '4', '1'],
]

const PPT_TABLE_TSV_MIME_TYPES = new Set([
  'text/tab-separated-values',
])
const PPT_TABLE_CSV_MIME_TYPES = new Set([
  'application/vnd.ms-excel',
  'text/comma-separated-values',
  'text/csv',
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
  return Array.from(files ?? [])
    .find((file) => isPPTTableCsvFile(file) || isPPTTableTsvFile(file)) ?? null
}

export function getPPTTableFileFromDataTransfer(dataTransfer: DataTransfer | null) {
  return getPPTTableFileFromList(dataTransfer?.files ?? null)
}

export function getPPTTableSourceFromDataTransfer(dataTransfer: DataTransfer | null) {
  if (!dataTransfer) {
    return null
  }

  const tabSeparatedText = dataTransfer.getData('text/tab-separated-values')

  if (tabSeparatedText) {
    return getPPTTableSourceFromText(tabSeparatedText, {
      format: 'text-tsv',
    })
  }

  const csvText = dataTransfer.getData('text/csv')

  if (csvText) {
    return getPPTTableSourceFromText(csvText, {
      format: 'canvas-csv',
    })
  }

  const htmlSource = getPPTTableSourceFromHTML(dataTransfer.getData('text/html'))

  if (htmlSource) {
    return htmlSource
  }

  return getPPTTableSourceFromText(dataTransfer.getData('text/plain'), {
    format: 'text-delimited',
  })
}

export async function readPPTTableFileSource(file: Blob & { name?: string }) {
  if (!isPPTTableCsvFile(file) && !isPPTTableTsvFile(file)) {
    return null
  }

  return getPPTTableSourceFromText(await readPPTBlobAsText(file), {
    format: isPPTTableTsvFile(file) ? 'text-tsv' : 'canvas-csv',
    name: file.name,
  })
}

export function getPPTTableSourceFromText(
  text: string,
  options: { format?: PPTTableImportFormat; name?: string } = {},
): PPTTableImportSource | null {
  if (!text.trim()) {
    return null
  }

  const parsedRows = parsePPTTableTextRows(text)

  if (!isPPTTableImportRows(parsedRows)) {
    return null
  }

  const rows = normalizePPTTableRows(parsedRows)

  return {
    ...(options.format === undefined ? {} : { format: options.format }),
    ...(options.name === undefined ? {} : { name: getPPTTableImportName(options.name) }),
    rows,
  }
}

export function getPPTTableSourceFromHTML(value: string) {
  if (!value || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(value, 'text/html')
  const table = doc.querySelector('table')

  if (!table) {
    return null
  }

  const parsedRows = parsePPTTableHTMLRows(table)

  if (!isPPTTableImportRows(parsedRows)) {
    return null
  }

  const rows = normalizePPTTableRows(parsedRows)

  return {
    format: 'text-html' as const,
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

function isPPTTableCsvFile(file: Blob & { name?: string }) {
  const mimeType = file.type.toLowerCase()
  const name = file.name?.toLowerCase() ?? ''

  return PPT_TABLE_CSV_MIME_TYPES.has(mimeType) || name.endsWith('.csv')
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

function parsePPTTableHTMLRows(table: Element) {
  const rows: string[][] = []
  const tableRows = Array.from(
    table instanceof HTMLTableElement
      ? table.rows
      : table.querySelectorAll(':scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr, :scope > tr'),
  )

  for (const [rowIndex, tableRow] of tableRows.entries()) {
    const row = rows[rowIndex] ?? []
    const cells = Array.from(tableRow.children).filter((cell) =>
      cell.tagName.toLowerCase() === 'td' ||
      cell.tagName.toLowerCase() === 'th',
    )
    let columnIndex = 0

    rows[rowIndex] = row

    for (const cell of cells) {
      while (row[columnIndex] !== undefined) {
        columnIndex += 1
      }

      const text = getPPTTableHTMLCellText(cell)
      const columnSpan = getPPTTableHTMLSpan(cell.getAttribute('colspan'))
      const rowSpan = getPPTTableHTMLSpan(cell.getAttribute('rowspan'))

      for (let rowOffset = 0; rowOffset < rowSpan; rowOffset += 1) {
        const targetRowIndex = rowIndex + rowOffset
        const targetRow = rows[targetRowIndex] ?? []

        rows[targetRowIndex] = targetRow

        for (let columnOffset = 0; columnOffset < columnSpan; columnOffset += 1) {
          const targetColumnIndex = columnIndex + columnOffset

          targetRow[targetColumnIndex] = rowOffset === 0 && columnOffset === 0
            ? text
            : targetRow[targetColumnIndex] ?? ''
        }
      }

      columnIndex += columnSpan
    }
  }

  return rows
}

function getPPTTableHTMLSpan(value: string | null) {
  const span = Number.parseInt(value ?? '1', 10)

  return Number.isFinite(span)
    ? clamp(span, 1, PPT_TABLE_MAX_COLUMNS)
    : 1
}

function getPPTTableHTMLCellText(cell: Element) {
  const clone = cell.cloneNode(true)

  if (clone instanceof Element) {
    clone.querySelectorAll('script, style, noscript').forEach((node) => node.remove())

    return normalizePPTTableCell(clone.textContent ?? '')
  }

  return normalizePPTTableCell(cell.textContent ?? '')
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
