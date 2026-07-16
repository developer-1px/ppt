import type { PPTInspectorProps } from './PPTInspectorContract'
import { PPTInspectorShell } from './PPTInspectorShell'
import { PPTSelectionInspectorPanel } from './PPTSelectionInspectorPanel'

export function PPTInspector({ model, onAction }: PPTInspectorProps) {
  return (
    <PPTInspectorShell
      model={model}
      onAction={onAction}
      selectionPanel={
        <PPTSelectionInspectorPanel model={model} onAction={onAction} />
      }
    />
  )
}
