import {
  getPPTTableColumnCount,
  stringifyPPTTableRows,
} from './pptImportExtension/tableImport'
import type { PPTTable } from './pptModel'

type PPTTableInspectorModel = {
  columnCount: number
  rowCount: number
  rowsText: string
}

export function getPPTTableInspectorModel(
  table: PPTTable,
): PPTTableInspectorModel {
  return {
    columnCount: getPPTTableColumnCount(table.rows),
    rowCount: table.rows.length,
    rowsText: stringifyPPTTableRows(table.rows),
  }
}
