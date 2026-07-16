import { PPTCommentInspectorFields } from './PPTCommentInspectorFields'
import { PPTImageInspectorFields } from './PPTImageInspectorFields'
import { PPTLineInspectorFields } from './PPTLineInspectorFields'
import { PPTObjectAnimationInspectorFields } from './PPTObjectAnimationInspectorFields'
import { PPTObjectPropertiesInspectorFields } from './PPTObjectPropertiesInspectorFields'
import type { PPTInspectorProps } from './PPTInspectorContract'
import { PPTShapeInspectorFields } from './PPTShapeInspectorFields'
import { PPTTableInspectorFields } from './PPTTableInspectorFields'
import { PPTTextInspectorFields } from './PPTTextInspectorFields'

export function PPTSelectionInspectorPanel({
  model,
  onAction,
}: PPTInspectorProps) {
  if (!model.selectedElement) {
    return <span className="ppt-muted">None</span>
  }

  return (
    <>
      <PPTObjectPropertiesInspectorFields model={model} onAction={onAction} />
      <PPTObjectAnimationInspectorFields model={model} onAction={onAction} />
      <PPTTextInspectorFields model={model} onAction={onAction} />
      <PPTShapeInspectorFields model={model} onAction={onAction} />
      <PPTImageInspectorFields model={model} onAction={onAction} />
      <PPTTableInspectorFields model={model} onAction={onAction} />
      <PPTCommentInspectorFields model={model} onAction={onAction} />
      <PPTLineInspectorFields model={model} onAction={onAction} />
    </>
  )
}
