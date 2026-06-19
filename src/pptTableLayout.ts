import type { PPTTable } from './pptModel'

export function getPPTTableCellFill(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return getPPTTableCellStyle(element, rowIndex, columnIndex)?.fill
}

export function getPPTTableCellColSpan(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return normalizePPTTableCellSpan(
    getPPTTableCellStyle(element, rowIndex, columnIndex)?.colSpan,
  )
}

export function getPPTTableCellBorders(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return getPPTTableCellStyle(element, rowIndex, columnIndex)?.borders
}

export function getPPTTableCellRowSpan(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return normalizePPTTableCellSpan(
    getPPTTableCellStyle(element, rowIndex, columnIndex)?.rowSpan,
  )
}

export function getPPTTableCellStyle(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return element.cellStyles?.[rowIndex]?.[columnIndex]
}

export function isPPTTableCellHidden(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return getPPTTableCellStyle(element, rowIndex, columnIndex)?.hidden === true
}

export function getPPTTableCellTextStyle(
  element: PPTTable,
  rowIndex: number,
  columnIndex: number,
) {
  return getPPTTableCellStyle(element, rowIndex, columnIndex)?.textStyle
}

function normalizePPTTableCellSpan(value: number | undefined) {
  return Number.isFinite(value) && value !== undefined && value > 1
    ? Math.floor(value)
    : 1
}

export function getPPTTableResolvedColumnWidths(element: PPTTable) {
  return resolvePPTTableTrackSizes({
    totalSize: element.geometry.w,
    trackCount: getPPTTableColumnCount(element.rows),
    trackSizes: element.columnWidths,
  })
}

export function getPPTTableResolvedRowHeights(element: PPTTable) {
  return resolvePPTTableTrackSizes({
    totalSize: element.geometry.h,
    trackCount: element.rows.length,
    trackSizes: element.rowHeights,
  })
}

export function getPPTTableColumnCount(rows: readonly (readonly string[])[]) {
  return Math.max(0, ...rows.map((row) => row.length))
}

function resolvePPTTableTrackSizes({
  totalSize,
  trackCount,
  trackSizes,
}: {
  totalSize: number
  trackCount: number
  trackSizes: readonly number[] | undefined
}) {
  if (trackCount <= 0) {
    return []
  }

  const fallbackSize = totalSize / trackCount

  if (!trackSizes || trackSizes.length < trackCount) {
    return Array.from({ length: trackCount }, () => fallbackSize)
  }

  const sizes = trackSizes.slice(0, trackCount)

  if (sizes.some((size) => !Number.isFinite(size) || size <= 0)) {
    return Array.from({ length: trackCount }, () => fallbackSize)
  }

  const sizeTotal = sizes.reduce((sum, size) => sum + size, 0)

  if (sizeTotal <= 0) {
    return Array.from({ length: trackCount }, () => fallbackSize)
  }

  const scale = totalSize / sizeTotal

  return sizes.map((size) => size * scale)
}
