import {
  formatPPTElementAnimationTrigger,
  formatPPTElementAnimationType,
  getPPTObjectAnimationDescriptor,
  isPPTElementAnimationTrigger,
  isPPTElementAnimationType,
  parsePPTElementAnimationOrder,
  parsePPTElementAnimationTime,
  PPT_ELEMENT_ANIMATION_TRIGGERS,
  PPT_ELEMENT_ANIMATION_TYPES,
  PPT_OBJECT_ANIMATION_LIMITS,
  PPT_OBJECT_ANIMATION_PACKAGE_TRIGGERS,
  PPT_OBJECT_ANIMATION_PACKAGE_TYPES,
} from '../../pptObjectAnimationAdapter'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTObjectAnimationInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTObjectAnimationInspectorFields({
  model,
  onAction,
}: PPTObjectAnimationInspectorFieldsProps) {
  const element = model.selectedElement
  const animation = model.selectedElementAnimation
  const { onElementAnimationChange } =
    createPPTInspectorActionDispatcher(onAction)

  if (!element || !animation) {
    return null
  }

  const descriptor = getPPTObjectAnimationDescriptor(model.slide, element)

  return (
    <div
      className="ppt-animation-grid"
      data-ppt-object-animation-inspector
      data-ppt-animation-delay={descriptor.delayMs ?? animation.delayMs}
      data-ppt-animation-duration={descriptor.durationMs ?? animation.durationMs}
      data-ppt-animation-limit-delay-max={PPT_OBJECT_ANIMATION_LIMITS.maxDelayMs}
      data-ppt-animation-limit-duration-max={PPT_OBJECT_ANIMATION_LIMITS.maxDurationMs}
      data-ppt-animation-limit-order-max={PPT_OBJECT_ANIMATION_LIMITS.maxBuildOrder}
      data-ppt-animation-model="slide-edit-object-animation"
      data-ppt-animation-order={descriptor.order ?? animation.order}
      data-ppt-animation-package-trigger={descriptor.trigger}
      data-ppt-animation-package-type={descriptor.type}
      data-ppt-animation-trigger={animation.trigger}
      data-ppt-animation-trigger-options={PPT_OBJECT_ANIMATION_PACKAGE_TRIGGERS.join(' ')}
      data-ppt-animation-type={animation.type}
      data-ppt-animation-type-options={PPT_OBJECT_ANIMATION_PACKAGE_TYPES.join(' ')}
    >
      <label className="ppt-field">
        <span>Animation</span>
        <select
          data-ppt-animation-command="update-object-animation"
          data-ppt-animation-field="type"
          data-ppt-animation-package-value={descriptor.type}
          value={animation.type}
          onChange={(event) => {
            if (isPPTElementAnimationType(event.target.value)) {
              onElementAnimationChange(
                element.id,
                'type',
                event.target.value,
              )
            }
          }}
        >
          {PPT_ELEMENT_ANIMATION_TYPES.map((type) => (
            <option key={type} value={type}>
              {formatPPTElementAnimationType(type)}
            </option>
          ))}
        </select>
      </label>
      <label className="ppt-field">
        <span>Trigger</span>
        <select
          data-ppt-animation-command="update-object-animation"
          data-ppt-animation-field="trigger"
          data-ppt-animation-package-value={descriptor.trigger}
          value={animation.trigger}
          onChange={(event) => {
            if (isPPTElementAnimationTrigger(event.target.value)) {
              onElementAnimationChange(
                element.id,
                'trigger',
                event.target.value,
              )
            }
          }}
        >
          {PPT_ELEMENT_ANIMATION_TRIGGERS.map((trigger) => (
            <option key={trigger} value={trigger}>
              {formatPPTElementAnimationTrigger(trigger)}
            </option>
          ))}
        </select>
      </label>
      <label className="ppt-field">
        <span>Duration</span>
        <input
          data-ppt-animation-command="update-object-animation"
          data-ppt-animation-field="durationMs"
          max={PPT_OBJECT_ANIMATION_LIMITS.maxDurationMs}
          min={PPT_OBJECT_ANIMATION_LIMITS.minDurationMs}
          step={100}
          type="number"
          value={descriptor.durationMs ?? animation.durationMs}
          onChange={(event) =>
            onElementAnimationChange(
              element.id,
              'durationMs',
              parsePPTElementAnimationTime(event.target.value),
            )}
        />
      </label>
      <label className="ppt-field">
        <span>Delay</span>
        <input
          data-ppt-animation-command="update-object-animation"
          data-ppt-animation-field="delayMs"
          max={PPT_OBJECT_ANIMATION_LIMITS.maxDelayMs}
          min={PPT_OBJECT_ANIMATION_LIMITS.minDelayMs}
          step={100}
          type="number"
          value={descriptor.delayMs ?? animation.delayMs}
          onChange={(event) =>
            onElementAnimationChange(
              element.id,
              'delayMs',
              parsePPTElementAnimationTime(event.target.value),
            )}
        />
      </label>
      <label className="ppt-field">
        <span>Order</span>
        <input
          data-ppt-animation-command="update-object-animation"
          data-ppt-animation-field="order"
          max={PPT_OBJECT_ANIMATION_LIMITS.maxBuildOrder}
          min={Math.max(1, PPT_OBJECT_ANIMATION_LIMITS.minBuildOrder)}
          step={1}
          type="number"
          value={descriptor.order ?? animation.order}
          onChange={(event) =>
            onElementAnimationChange(
              element.id,
              'order',
              parsePPTElementAnimationOrder(
                event.target.value,
                model.slide.elements.length,
              ),
            )}
        />
      </label>
    </div>
  )
}
