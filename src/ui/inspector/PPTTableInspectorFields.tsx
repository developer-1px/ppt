import { getPPTTableInspectorModel } from '../../pptTableAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTTableInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTTableInspectorFields({
  model,
  onAction,
}: PPTTableInspectorFieldsProps) {
  const element = model.selectedElement

  if (element?.kind !== 'table') {
    return null
  }

  const { onTableRowsChange } = createPPTInspectorActionDispatcher(onAction)
  const tableModel = getPPTTableInspectorModel(element)

  return (
    <>
      <label className="ppt-field">
        <span>Rows</span>
        <textarea
          data-ppt-style-field="table-data"
          value={tableModel.rowsText}
          onChange={(event) =>
            onTableRowsChange(element.id, event.target.value)}
        />
      </label>
      <span className="ppt-muted" data-ppt-table-inspector-size>
        {tableModel.rowCount} x {tableModel.columnCount}
      </span>
    </>
  )
}
