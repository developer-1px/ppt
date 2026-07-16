import {
  getPPTLineMarker,
  getPPTLineRoute,
  isPPTLineMarker,
  isPPTLineRoute,
  PPT_LINE_MARKER_OPTIONS,
  PPT_LINE_ROUTE_OPTIONS,
} from '../../pptLineAdapter'
import {
  getPPTStrokeDash,
  getPPTStrokeLineStyleDescriptor,
  isPPTStrokeDash,
  PPT_STROKE_DASH_OPTIONS,
} from '../../pptObjectAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'
import { PPTColorSwatchStrip } from './PPTColorSwatchStrip'

type PPTLineInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTLineInspectorFields({
  model,
  onAction,
}: PPTLineInspectorFieldsProps) {
  const element = model.selectedElement

  if (element?.kind !== 'line' && element?.kind !== 'freeform') {
    return null
  }

  const {
    onElementStrokeChange,
    onLineMarkerChange,
    onLineRouteChange,
  } = createPPTInspectorActionDispatcher(onAction)
  const strokeLineStyleDescriptor = getPPTStrokeLineStyleDescriptor(
    model.slide.id,
    element,
  )

  return (
    <>
      <div className="ppt-geometry-grid">
        <div
          className="ppt-color-control"
          data-ppt-color-control="line-stroke"
        >
          <label className="ppt-field">
            <span>Stroke</span>
            <input
              data-ppt-style-field="line-stroke-color"
              type="color"
              value={element.stroke.color}
              onChange={(event) =>
                onElementStrokeChange(
                  element.id,
                  'color',
                  event.target.value,
                )}
            />
          </label>
          <PPTColorSwatchStrip
            model={model}
            onAction={onAction}
            target={{
              channel: 'line-stroke',
              color: element.stroke.color,
              elementId: element.id,
            }}
          />
        </div>
        <label className="ppt-field">
          <span>Width</span>
          <input
            data-ppt-style-field="line-stroke-width"
            type="number"
            value={element.stroke.width}
            onChange={(event) =>
              onElementStrokeChange(
                element.id,
                'width',
                Number(event.target.value),
              )}
          />
        </label>
      </div>
      <label className="ppt-field">
        <span>Dash</span>
        <select
          data-ppt-style-field="line-stroke-dash"
          data-ppt-stroke-line-style-attribute={
            strokeLineStyleDescriptor.metadata.attribute
          }
          data-ppt-stroke-line-style-attribute-value={
            strokeLineStyleDescriptor.metadata.attributeValue
          }
          data-ppt-stroke-line-style-command={
            strokeLineStyleDescriptor.field.commandId
          }
          data-ppt-stroke-line-style-control={
            strokeLineStyleDescriptor.field.control
          }
          data-ppt-stroke-line-style-surface={
            strokeLineStyleDescriptor.surface
          }
          value={
            strokeLineStyleDescriptor.value ?? getPPTStrokeDash(element.stroke)
          }
          onChange={(event) => {
            if (isPPTStrokeDash(event.target.value)) {
              onElementStrokeChange(element.id, 'dash', event.target.value)
            }
          }}
        >
          {(strokeLineStyleDescriptor.field.options ??
            PPT_STROKE_DASH_OPTIONS).map((option) => (
            <option key={option.id} value={option.id}>
              {option.label}
            </option>
          ))}
        </select>
      </label>
      {element.kind === 'line' ? (
        <>
          <label className="ppt-field">
            <span>Route</span>
            <select
              data-ppt-style-field="line-route"
              value={getPPTLineRoute(element)}
              onChange={(event) => {
                if (isPPTLineRoute(event.target.value)) {
                  onLineRouteChange(element.id, event.target.value)
                }
              }}
            >
              {PPT_LINE_ROUTE_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
          <div className="ppt-geometry-grid">
            {(['startMarker', 'endMarker'] as const).map((field) => (
              <label className="ppt-field" key={field}>
                <span>{field === 'startMarker' ? 'Start' : 'End'}</span>
                <select
                  data-ppt-style-field={
                    field === 'startMarker'
                      ? 'line-start-marker'
                      : 'line-end-marker'
                  }
                  value={getPPTLineMarker(element, field)}
                  onChange={(event) => {
                    if (isPPTLineMarker(event.target.value)) {
                      onLineMarkerChange(
                        element.id,
                        field,
                        event.target.value,
                      )
                    }
                  }}
                >
                  {PPT_LINE_MARKER_OPTIONS.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
        </>
      ) : null}
    </>
  )
}
