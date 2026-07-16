import {
  getPPTCornerRadiusDescriptor,
  getPPTFillOpacity,
  getPPTFillOpacityDescriptor,
  getPPTShapeCornerRadius,
  getPPTStrokeDash,
  getPPTStrokeLineStyleDescriptor,
  isPPTShapeKind,
  isPPTStrokeDash,
  parsePPTFillOpacity,
  parsePPTShapeCornerRadius,
  PPT_FILL_OPACITY_MAX,
  PPT_FILL_OPACITY_MIN,
  PPT_FILL_OPACITY_STEP,
  PPT_SHAPE_CORNER_RADIUS_MAX,
  PPT_SHAPE_CORNER_RADIUS_MIN,
  PPT_SHAPE_CORNER_RADIUS_STEP,
  PPT_STROKE_DASH_OPTIONS,
} from '../../pptObjectAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'
import { PPTColorSwatchStrip } from './PPTColorSwatchStrip'

type PPTShapeInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTShapeInspectorFields({
  model,
  onAction,
}: PPTShapeInspectorFieldsProps) {
  const element = model.selectedElement

  if (element?.kind !== 'shape') {
    return null
  }

  const {
    onElementStrokeChange,
    onShapeCornerRadiusChange,
    onShapeFillChange,
    onShapeKindChange,
  } = createPPTInspectorActionDispatcher(onAction)
  const cornerRadiusDescriptor = getPPTCornerRadiusDescriptor(
    model.slide.id,
    element,
  )
  const fillOpacityDescriptor = getPPTFillOpacityDescriptor(
    model.slide.id,
    element,
  )
  const strokeLineStyleDescriptor = getPPTStrokeLineStyleDescriptor(
    model.slide.id,
    element,
  )

  return (
    <>
      <label className="ppt-field">
        <span>Shape</span>
        <select
          data-ppt-style-field="shape"
          value={element.shape}
          onChange={(event) => {
            if (isPPTShapeKind(event.target.value)) {
              onShapeKindChange(element.id, event.target.value)
            }
          }}
        >
          <option value="rect">Rectangle</option>
          <option value="ellipse">Oval</option>
          <option value="diamond">Diamond</option>
        </select>
      </label>
      {element.shape === 'rect' ? (
        <label className="ppt-field">
          <span>Corner radius</span>
          <input
            data-ppt-style-field="shape-corner-radius"
            data-ppt-corner-radius-attribute={
              cornerRadiusDescriptor.metadata.attribute
            }
            data-ppt-corner-radius-attribute-value={
              cornerRadiusDescriptor.metadata.attributeValue
            }
            data-ppt-corner-radius-command={
              cornerRadiusDescriptor.field.commandId
            }
            data-ppt-corner-radius-control={
              cornerRadiusDescriptor.field.control
            }
            data-ppt-corner-radius-supported={
              cornerRadiusDescriptor.isSupported ? 'true' : 'false'
            }
            data-ppt-corner-radius-surface={cornerRadiusDescriptor.surface}
            max={
              cornerRadiusDescriptor.field.max ??
              PPT_SHAPE_CORNER_RADIUS_MAX
            }
            min={
              cornerRadiusDescriptor.field.min ??
              PPT_SHAPE_CORNER_RADIUS_MIN
            }
            step={
              cornerRadiusDescriptor.field.step ??
              PPT_SHAPE_CORNER_RADIUS_STEP
            }
            type="number"
            value={
              cornerRadiusDescriptor.value ?? getPPTShapeCornerRadius(element)
            }
            onChange={(event) =>
              onShapeCornerRadiusChange(
                element.id,
                parsePPTShapeCornerRadius(event.target.value),
              )}
          />
        </label>
      ) : null}
      <div className="ppt-geometry-grid">
        <div className="ppt-color-control" data-ppt-color-control="shape-fill">
          <label className="ppt-field">
            <span>Fill</span>
            <input
              data-ppt-style-field="fill"
              type="color"
              value={element.fill.color}
              onChange={(event) =>
                onShapeFillChange(element.id, 'color', event.target.value)}
            />
          </label>
          <PPTColorSwatchStrip
            model={model}
            onAction={onAction}
            target={{
              channel: 'shape-fill',
              color: element.fill.color,
              elementId: element.id,
            }}
          />
        </div>
        <div
          className="ppt-color-control"
          data-ppt-color-control="shape-stroke"
        >
          <label className="ppt-field">
            <span>Stroke</span>
            <input
              data-ppt-style-field="stroke-color"
              type="color"
              value={element.stroke?.color ?? '#111827'}
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
              channel: 'shape-stroke',
              color: element.stroke?.color ?? '#111827',
              elementId: element.id,
            }}
          />
        </div>
      </div>
      <label className="ppt-field">
        <span>Fill opacity</span>
        <input
          data-ppt-style-field="fill-opacity"
          data-ppt-fill-opacity-attribute={
            fillOpacityDescriptor.metadata.attribute
          }
          data-ppt-fill-opacity-attribute-value={
            fillOpacityDescriptor.metadata.attributeValue
          }
          data-ppt-fill-opacity-command={fillOpacityDescriptor.field.commandId}
          data-ppt-fill-opacity-control={fillOpacityDescriptor.field.control}
          data-ppt-fill-opacity-surface={fillOpacityDescriptor.surface}
          max={fillOpacityDescriptor.field.max ?? PPT_FILL_OPACITY_MAX}
          min={fillOpacityDescriptor.field.min ?? PPT_FILL_OPACITY_MIN}
          step={fillOpacityDescriptor.field.step ?? PPT_FILL_OPACITY_STEP}
          type="number"
          value={
            fillOpacityDescriptor.value ?? getPPTFillOpacity(element.fill)
          }
          onChange={(event) =>
            onShapeFillChange(
              element.id,
              'opacity',
              parsePPTFillOpacity(event.target.value),
            )}
        />
      </label>
      <label className="ppt-field">
        <span>Stroke width</span>
        <input
          data-ppt-style-field="stroke-width"
          type="number"
          value={element.stroke?.width ?? 0}
          onChange={(event) =>
            onElementStrokeChange(
              element.id,
              'width',
              Number(event.target.value),
            )}
        />
      </label>
      <label className="ppt-field">
        <span>Dash</span>
        <select
          data-ppt-style-field="stroke-dash"
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
    </>
  )
}
