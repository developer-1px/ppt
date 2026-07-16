import {
  getPPTElementAltText,
  getPPTElementHyperlink,
  getPPTElementOpacity,
  getPPTElementShadow,
  getPPTObjectAccessibilityDescriptor,
  getPPTObjectHyperlinkDescriptor,
  getPPTObjectOpacityDescriptor,
  getPPTObjectShadowDescriptor,
  getPPTObjectShadowField,
  hasPPTElementShadow,
  parsePPTElementOpacity,
  parsePPTElementShadowAngle,
  parsePPTElementShadowBlur,
  parsePPTElementShadowDistance,
  parsePPTElementShadowOpacity,
  PPT_ALT_TEXT_MAX_LENGTH,
  PPT_ELEMENT_OPACITY_MAX,
  PPT_ELEMENT_OPACITY_MIN,
  PPT_ELEMENT_OPACITY_STEP,
  PPT_ELEMENT_SHADOW_ANGLE_MAX,
  PPT_ELEMENT_SHADOW_ANGLE_MIN,
  PPT_ELEMENT_SHADOW_BLUR_MAX,
  PPT_ELEMENT_SHADOW_DISTANCE_MAX,
  PPT_ELEMENT_SHADOW_OPACITY_MAX,
  PPT_ELEMENT_SHADOW_OPACITY_MIN,
  PPT_ELEMENT_SHADOW_OPACITY_STEP,
  PPT_HYPERLINK_URL_MAX_LENGTH,
} from '../../pptObjectAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTObjectPropertiesInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTObjectPropertiesInspectorFields({
  model,
  onAction,
}: PPTObjectPropertiesInspectorFieldsProps) {
  const element = model.selectedElement
  const {
    onElementAltTextChange,
    onElementGeometryChange,
    onElementHyperlinkChange,
    onElementNameChange,
    onElementOpacityChange,
    onElementRotationChange,
    onElementShadowChange,
  } = createPPTInspectorActionDispatcher(onAction)

  if (!element) {
    return null
  }

  const objectOpacityDescriptor = getPPTObjectOpacityDescriptor(
    model.slide.id,
    element,
  )
  const objectHyperlinkDescriptor = getPPTObjectHyperlinkDescriptor(
    model.slide.id,
    element,
  )
  const objectHyperlinkUrlField = objectHyperlinkDescriptor.fields.find(
    (field) => field.id === 'url',
  )
  const objectAccessibilityDescriptor = getPPTObjectAccessibilityDescriptor(
    model.slide.id,
    element,
  )
  const objectAccessibilityAltTextField =
    objectAccessibilityDescriptor.fields.find((field) => field.id === 'altText')
  const objectShadowDescriptor = getPPTObjectShadowDescriptor(
    model.slide.id,
    element,
  )
  const objectShadowEnabledField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'enabled',
  )
  const objectShadowColorField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'color',
  )
  const objectShadowOpacityField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'opacity',
  )
  const objectShadowBlurField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'blur',
  )
  const objectShadowDistanceField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'distance',
  )
  const objectShadowAngleField = getPPTObjectShadowField(
    objectShadowDescriptor,
    'angle',
  )
  const elementHyperlink = getPPTElementHyperlink(element)
  const elementAltText = getPPTElementAltText(element) ?? ''
  const elementShadow = getPPTElementShadow(element)
  const elementShadowEnabled = hasPPTElementShadow(element)
  const objectShadowEnabled =
    objectShadowDescriptor.metadata.isEnabled ?? elementShadowEnabled

  return (
    <>
      <label className="ppt-field">
        <span>Name</span>
        <input
          data-ppt-style-field="name"
          value={element.name}
          onChange={(event) =>
            onElementNameChange(element.id, event.target.value)}
        />
      </label>
      <label className="ppt-field">
        <span>Opacity</span>
        <input
          data-ppt-style-field="opacity"
          data-ppt-object-opacity-attribute={objectOpacityDescriptor.metadata.attribute}
          data-ppt-object-opacity-attribute-value={objectOpacityDescriptor.metadata.attributeValue}
          data-ppt-object-opacity-command={objectOpacityDescriptor.field.commandId}
          data-ppt-object-opacity-control={objectOpacityDescriptor.field.control}
          data-ppt-object-opacity-surface={objectOpacityDescriptor.surface}
          max={objectOpacityDescriptor.field.max ?? PPT_ELEMENT_OPACITY_MAX}
          min={objectOpacityDescriptor.field.min ?? PPT_ELEMENT_OPACITY_MIN}
          step={objectOpacityDescriptor.field.step ?? PPT_ELEMENT_OPACITY_STEP}
          type="number"
          value={objectOpacityDescriptor.value ?? getPPTElementOpacity(element)}
          onChange={(event) =>
            onElementOpacityChange(
              element.id,
              parsePPTElementOpacity(event.target.value),
            )}
        />
      </label>
      <label className="ppt-field">
        <span>Link</span>
        <input
          data-ppt-style-field="hyperlink"
          data-ppt-hyperlink-attribute={objectHyperlinkDescriptor.metadata.attribute}
          data-ppt-hyperlink-attribute-value={objectHyperlinkDescriptor.metadata.attributeValue}
          data-ppt-hyperlink-command={objectHyperlinkUrlField?.commandId}
          data-ppt-hyperlink-control={objectHyperlinkUrlField?.control}
          data-ppt-hyperlink-enabled={objectHyperlinkDescriptor.metadata.isEnabled ? 'true' : 'false'}
          data-ppt-hyperlink-surface={objectHyperlinkDescriptor.surface}
          data-ppt-hyperlink-validation={objectHyperlinkDescriptor.metadata.validation.reason}
          maxLength={PPT_HYPERLINK_URL_MAX_LENGTH}
          placeholder="https://example.com"
          value={objectHyperlinkDescriptor.hyperlink.url ?? elementHyperlink?.url ?? ''}
          onChange={(event) =>
            onElementHyperlinkChange(element.id, event.target.value)}
        />
      </label>
      <label className="ppt-field">
        <span>Alt text</span>
        <textarea
          data-ppt-style-field="alt-text"
          data-ppt-accessibility-attribute={objectAccessibilityDescriptor.metadata.attribute}
          data-ppt-accessibility-attribute-value={objectAccessibilityDescriptor.metadata.attributeValue}
          data-ppt-accessibility-command={objectAccessibilityAltTextField?.commandId}
          data-ppt-accessibility-control={objectAccessibilityAltTextField?.control}
          data-ppt-accessibility-described={objectAccessibilityDescriptor.metadata.isDescribed ? 'true' : 'false'}
          data-ppt-accessibility-surface={objectAccessibilityDescriptor.surface}
          maxLength={PPT_ALT_TEXT_MAX_LENGTH}
          value={objectAccessibilityDescriptor.value.altText ?? elementAltText}
          onChange={(event) =>
            onElementAltTextChange(element.id, event.target.value)}
        />
      </label>
      <label className="ppt-checkbox-field">
        <input
          checked={objectShadowEnabled}
          data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
          data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
          data-ppt-shadow-command={objectShadowEnabledField?.commandId}
          data-ppt-shadow-control={objectShadowEnabledField?.control}
          data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
          data-ppt-shadow-field="enabled"
          data-ppt-shadow-surface={objectShadowDescriptor.surface}
          type="checkbox"
          onChange={(event) =>
            onElementShadowChange(element.id, 'enabled', event.target.checked)}
        />
        <span>Shadow</span>
      </label>
      <div
        className="ppt-geometry-grid"
        data-ppt-shadow-angle={objectShadowDescriptor.shadow.angle ?? elementShadow.angle}
        data-ppt-shadow-blur={objectShadowDescriptor.shadow.blur ?? elementShadow.blur}
        data-ppt-shadow-color={objectShadowDescriptor.shadow.color ?? elementShadow.color}
        data-ppt-shadow-distance={objectShadowDescriptor.shadow.distance ?? elementShadow.distance}
        data-ppt-shadow-enabled={objectShadowEnabled ? 'true' : 'false'}
        data-ppt-shadow-inspector
        data-ppt-shadow-opacity={objectShadowDescriptor.shadow.opacity ?? elementShadow.opacity}
      >
        <label className="ppt-field">
          <span>Color</span>
          <input
            data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
            data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
            data-ppt-shadow-command={objectShadowColorField?.commandId}
            data-ppt-shadow-control={objectShadowColorField?.control}
            data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
            data-ppt-shadow-field="color"
            data-ppt-shadow-surface={objectShadowDescriptor.surface}
            disabled={!objectShadowEnabled}
            type="color"
            value={objectShadowDescriptor.shadow.color ?? elementShadow.color}
            onChange={(event) =>
              onElementShadowChange(element.id, 'color', event.target.value)}
          />
        </label>
        <label className="ppt-field">
          <span>Shadow opacity</span>
          <input
            data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
            data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
            data-ppt-shadow-command={objectShadowOpacityField?.commandId}
            data-ppt-shadow-control={objectShadowOpacityField?.control}
            data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
            data-ppt-shadow-field="opacity"
            data-ppt-shadow-surface={objectShadowDescriptor.surface}
            data-ppt-shadow-unit={objectShadowOpacityField?.unit}
            disabled={!objectShadowEnabled}
            max={objectShadowOpacityField?.max ?? PPT_ELEMENT_SHADOW_OPACITY_MAX}
            min={objectShadowOpacityField?.min ?? PPT_ELEMENT_SHADOW_OPACITY_MIN}
            step={objectShadowOpacityField?.step ?? PPT_ELEMENT_SHADOW_OPACITY_STEP}
            type="number"
            value={objectShadowDescriptor.shadow.opacity ?? elementShadow.opacity}
            onChange={(event) =>
              onElementShadowChange(
                element.id,
                'opacity',
                parsePPTElementShadowOpacity(event.target.value),
              )}
          />
        </label>
        <label className="ppt-field">
          <span>Blur</span>
          <input
            data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
            data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
            data-ppt-shadow-command={objectShadowBlurField?.commandId}
            data-ppt-shadow-control={objectShadowBlurField?.control}
            data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
            data-ppt-shadow-field="blur"
            data-ppt-shadow-surface={objectShadowDescriptor.surface}
            data-ppt-shadow-unit={objectShadowBlurField?.unit}
            disabled={!objectShadowEnabled}
            max={objectShadowBlurField?.max ?? PPT_ELEMENT_SHADOW_BLUR_MAX}
            min={objectShadowBlurField?.min ?? 0}
            step={objectShadowBlurField?.step ?? 1}
            type="number"
            value={objectShadowDescriptor.shadow.blur ?? elementShadow.blur}
            onChange={(event) =>
              onElementShadowChange(
                element.id,
                'blur',
                parsePPTElementShadowBlur(event.target.value),
              )}
          />
        </label>
        <label className="ppt-field">
          <span>Distance</span>
          <input
            data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
            data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
            data-ppt-shadow-command={objectShadowDistanceField?.commandId}
            data-ppt-shadow-control={objectShadowDistanceField?.control}
            data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
            data-ppt-shadow-field="distance"
            data-ppt-shadow-surface={objectShadowDescriptor.surface}
            data-ppt-shadow-unit={objectShadowDistanceField?.unit}
            disabled={!objectShadowEnabled}
            max={objectShadowDistanceField?.max ?? PPT_ELEMENT_SHADOW_DISTANCE_MAX}
            min={objectShadowDistanceField?.min ?? 0}
            step={objectShadowDistanceField?.step ?? 1}
            type="number"
            value={objectShadowDescriptor.shadow.distance ?? elementShadow.distance}
            onChange={(event) =>
              onElementShadowChange(
                element.id,
                'distance',
                parsePPTElementShadowDistance(event.target.value),
              )}
          />
        </label>
        <label className="ppt-field">
          <span>Angle</span>
          <input
            data-ppt-shadow-attribute={objectShadowDescriptor.metadata.attribute}
            data-ppt-shadow-attribute-value={objectShadowDescriptor.metadata.attributeValue}
            data-ppt-shadow-command={objectShadowAngleField?.commandId}
            data-ppt-shadow-control={objectShadowAngleField?.control}
            data-ppt-shadow-descriptor-enabled={objectShadowDescriptor.metadata.isEnabled ? 'true' : 'false'}
            data-ppt-shadow-field="angle"
            data-ppt-shadow-surface={objectShadowDescriptor.surface}
            data-ppt-shadow-unit={objectShadowAngleField?.unit}
            disabled={!objectShadowEnabled}
            max={objectShadowAngleField?.max ?? PPT_ELEMENT_SHADOW_ANGLE_MAX}
            min={objectShadowAngleField?.min ?? PPT_ELEMENT_SHADOW_ANGLE_MIN}
            step={objectShadowAngleField?.step ?? 1}
            type="number"
            value={objectShadowDescriptor.shadow.angle ?? elementShadow.angle}
            onChange={(event) =>
              onElementShadowChange(
                element.id,
                'angle',
                parsePPTElementShadowAngle(event.target.value),
              )}
          />
        </label>
      </div>
      <div className="ppt-geometry-grid">
        {(['x', 'y', 'w', 'h'] as const).map((field) => (
          <label className="ppt-field" key={field}>
            <span>{field.toUpperCase()}</span>
            <input
              data-ppt-geometry-field={field}
              type="number"
              value={Math.round(element.geometry[field])}
              onChange={(event) =>
                onElementGeometryChange(
                  element.id,
                  field,
                  Number(event.target.value),
                )}
            />
          </label>
        ))}
        <label className="ppt-field">
          <span>ROT</span>
          <input
            data-ppt-geometry-field="rotation"
            type="number"
            value={Math.round(element.geometry.rotation ?? 0)}
            onChange={(event) =>
              onElementRotationChange(element.id, Number(event.target.value))}
          />
        </label>
      </div>
    </>
  )
}
