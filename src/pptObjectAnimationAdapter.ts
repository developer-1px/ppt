import { clampPPTCanvasValue } from './pptCanvasCoreAdapter'
import type {
  PPTElement,
  PPTElementAnimation,
  PPTSlide,
} from './pptModel'
import {
  createSlideEditObjectAnimationDescriptor,
  getSlideEditObjectAnimationBuildOrder,
  getSlideEditObjectAnimationCSSStyle,
  normalizeSlideEditObjectAnimationDelayMs,
  normalizeSlideEditObjectAnimationDurationMs,
  normalizeSlideEditObjectAnimationOrder,
  SLIDE_EDIT_OBJECT_ANIMATION_LIMITS,
  SLIDE_EDIT_OBJECT_ANIMATION_TRIGGERS,
  SLIDE_EDIT_OBJECT_ANIMATION_TYPES,
  type SlideEditBuiltInAnimationTrigger,
  type SlideEditBuiltInAnimationType,
  type SlideEditObjectAnimationDescriptor,
  type SlideEditObjectAnimationUpdateCommand,
} from './pptSlideEditAffordanceAdapter'

type PPTElementAnimationType = PPTElementAnimation['type']
type PPTElementAnimationTrigger = PPTElementAnimation['trigger']
export type PPTElementAnimationUpdateField = keyof PPTElementAnimation
type PPTElementAnimationCSSStyle = ReturnType<
  typeof getSlideEditObjectAnimationCSSStyle
>

export const PPT_ELEMENT_ANIMATION_TYPES = Object.freeze([
  'none',
  'fadeIn',
  'flyIn',
] as const satisfies readonly PPTElementAnimationType[])

export const PPT_ELEMENT_ANIMATION_TRIGGERS = Object.freeze([
  'onClick',
  'withPrevious',
] as const satisfies readonly PPTElementAnimationTrigger[])

const PPT_DEFAULT_ELEMENT_ANIMATION = Object.freeze({
  delayMs: 0,
  durationMs: 400,
  order: 1,
  trigger: 'onClick',
  type: 'none',
} as const satisfies PPTElementAnimation)

export const PPT_OBJECT_ANIMATION_LIMITS = SLIDE_EDIT_OBJECT_ANIMATION_LIMITS
export const PPT_OBJECT_ANIMATION_PACKAGE_TRIGGERS = Object.freeze(
  SLIDE_EDIT_OBJECT_ANIMATION_TRIGGERS.map((trigger) => trigger.id),
)
export const PPT_OBJECT_ANIMATION_PACKAGE_TYPES = Object.freeze(
  SLIDE_EDIT_OBJECT_ANIMATION_TYPES.map((type) => type.id),
)

export function getPPTElementAnimation(
  element: PPTElement,
  slide?: PPTSlide,
): PPTElementAnimation {
  return normalizePPTElementAnimation(
    element.animation ?? {
      ...PPT_DEFAULT_ELEMENT_ANIMATION,
      order: getPPTElementDefaultAnimationOrder(element.id, slide),
    },
    slide,
    element.id,
  )
}

export function normalizePPTElementAnimation(
  animation: PPTElementAnimation,
  slide?: PPTSlide,
  elementId?: string,
): PPTElementAnimation {
  return {
    delayMs: clampPPTElementAnimationTime(animation.delayMs),
    durationMs: clampPPTElementAnimationTime(animation.durationMs),
    order: clampPPTElementAnimationOrder(
      animation.order,
      slide?.elements.length,
      elementId
        ? getPPTElementDefaultAnimationOrder(elementId, slide)
        : undefined,
    ),
    trigger: isPPTElementAnimationTrigger(animation.trigger)
      ? animation.trigger
      : 'onClick',
    type: isPPTElementAnimationType(animation.type)
      ? animation.type
      : 'none',
  }
}

export function getPPTObjectAnimationDescriptor(
  slide: PPTSlide,
  element: PPTElement,
): SlideEditObjectAnimationDescriptor<
  string,
  string,
  SlideEditBuiltInAnimationType,
  SlideEditBuiltInAnimationTrigger
> {
  const animation = getPPTElementAnimation(element, slide)

  return createSlideEditObjectAnimationDescriptor({
    delayMs: animation.delayMs,
    durationMs: animation.durationMs,
    objectId: element.id,
    order: animation.order,
    slideId: slide.id,
    trigger: toSlideEditObjectAnimationTrigger(animation.trigger),
    type: toSlideEditObjectAnimationType(animation.type),
  })
}

export function getPPTSlideAnimationBuildOrder(
  slide: PPTSlide,
  options: { visibleOnly?: boolean } = {},
) {
  return getSlideEditObjectAnimationBuildOrder(
    slide.elements
      .filter((element) => !options.visibleOnly || element.visible !== false)
      .map((element) => getPPTObjectAnimationDescriptor(slide, element)),
  )
}

export function toSlideEditObjectAnimationCommand({
  elementId,
  field,
  slideId,
  value,
}: {
  elementId: string
  field: PPTElementAnimationUpdateField
  slideId: string
  value: PPTElementAnimation[PPTElementAnimationUpdateField]
}): SlideEditObjectAnimationUpdateCommand<string, string> {
  if (field === 'type') {
    return {
      fieldId: 'type',
      id: 'update-object-animation',
      objectId: elementId,
      slideId,
      value: toSlideEditObjectAnimationType(value as PPTElementAnimationType),
    }
  }

  if (field === 'trigger') {
    return {
      fieldId: 'trigger',
      id: 'update-object-animation',
      objectId: elementId,
      slideId,
      value: toSlideEditObjectAnimationTrigger(
        value as PPTElementAnimationTrigger,
      ),
    }
  }

  switch (field) {
    case 'delayMs':
      return {
        fieldId: 'delayMs',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
    case 'durationMs':
      return {
        fieldId: 'durationMs',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
    case 'order':
      return {
        fieldId: 'order',
        id: 'update-object-animation',
        objectId: elementId,
        slideId,
        value: Number(value),
      }
  }
}

export function toPPTElementAnimationUpdate(
  command: SlideEditObjectAnimationUpdateCommand<string, string>,
): {
  field: PPTElementAnimationUpdateField
  value: PPTElementAnimation[PPTElementAnimationUpdateField]
} {
  if (command.fieldId === 'type') {
    return {
      field: 'type',
      value: toPPTElementAnimationType(command.value),
    }
  }

  if (command.fieldId === 'trigger') {
    return {
      field: 'trigger',
      value: toPPTElementAnimationTrigger(command.value),
    }
  }

  switch (command.fieldId) {
    case 'delayMs':
      return {
        field: 'delayMs',
        value: normalizeSlideEditObjectAnimationDelayMs(command.value),
      }
    case 'durationMs':
      return {
        field: 'durationMs',
        value: normalizeSlideEditObjectAnimationDurationMs(command.value),
      }
    case 'order':
      return {
        field: 'order',
        value: normalizeSlideEditObjectAnimationOrder(command.value),
      }
  }
}

function toSlideEditObjectAnimationType(
  type: PPTElementAnimationType,
): SlideEditBuiltInAnimationType {
  switch (type) {
    case 'fadeIn':
      return 'fade-in'
    case 'flyIn':
      return 'fly-in'
    case 'none':
      return 'none'
  }
}

function toPPTElementAnimationType(
  type: string,
): PPTElementAnimationType {
  switch (type) {
    case 'fade-in':
      return 'fadeIn'
    case 'fly-in':
      return 'flyIn'
    default:
      return 'none'
  }
}

function toSlideEditObjectAnimationTrigger(
  trigger: PPTElementAnimationTrigger,
): SlideEditBuiltInAnimationTrigger {
  switch (trigger) {
    case 'onClick':
      return 'on-click'
    case 'withPrevious':
      return 'with-previous'
  }
}

function toPPTElementAnimationTrigger(
  trigger: string,
): PPTElementAnimationTrigger {
  return trigger === 'with-previous' ? 'withPrevious' : 'onClick'
}

function getPPTElementDefaultAnimationOrder(
  elementId: string,
  slide?: PPTSlide,
) {
  const index = slide?.elements.findIndex(
    (element) => element.id === elementId,
  ) ?? -1

  return index >= 0 ? index + 1 : PPT_DEFAULT_ELEMENT_ANIMATION.order
}

function clampPPTElementAnimationTime(value: number) {
  return normalizeSlideEditObjectAnimationDurationMs(value)
}

function clampPPTElementAnimationOrder(
  value: number,
  elementCount: number = Number.MAX_SAFE_INTEGER,
  fallback: number = PPT_DEFAULT_ELEMENT_ANIMATION.order,
) {
  return clampPPTCanvasValue(
    Number.isFinite(value) ? Math.round(value) : fallback,
    1,
    Math.max(1, elementCount),
  )
}

export function parsePPTElementAnimationTime(value: string) {
  return clampPPTElementAnimationTime(Number(value))
}

export function parsePPTElementAnimationOrder(
  value: string,
  elementCount: number,
) {
  return clampPPTElementAnimationOrder(
    normalizeSlideEditObjectAnimationOrder(Number(value)),
    elementCount,
  )
}

export function getPPTElementAnimationStyle(
  animation: PPTElementAnimation,
): PPTElementAnimationCSSStyle {
  return getSlideEditObjectAnimationCSSStyle({
    delayMs: animation.delayMs,
    durationMs: animation.durationMs,
  })
}

export function formatPPTElementAnimationType(
  type: PPTElementAnimationType,
) {
  switch (type) {
    case 'fadeIn':
      return 'Fade in'
    case 'flyIn':
      return 'Fly in'
    case 'none':
      return 'None'
  }
}

export function formatPPTElementAnimationTrigger(
  trigger: PPTElementAnimationTrigger,
) {
  switch (trigger) {
    case 'onClick':
      return 'On click'
    case 'withPrevious':
      return 'With previous'
  }
}

export function isPPTElementAnimationType(
  value: string,
): value is PPTElementAnimationType {
  return (PPT_ELEMENT_ANIMATION_TYPES as readonly string[]).includes(value)
}

export function isPPTElementAnimationTrigger(
  value: string,
): value is PPTElementAnimationTrigger {
  return (PPT_ELEMENT_ANIMATION_TRIGGERS as readonly string[]).includes(value)
}

export function getPPTElementAnimationTypeFromJSONValue(
  value: unknown,
): PPTElementAnimationType | undefined {
  const type = typeof value === 'string' ? value.trim() : ''

  return isPPTElementAnimationType(type) ? type : undefined
}

export function getPPTElementAnimationTriggerFromJSONValue(
  value: unknown,
): PPTElementAnimationTrigger | undefined {
  const trigger = typeof value === 'string' ? value.trim() : ''

  return isPPTElementAnimationTrigger(trigger) ? trigger : undefined
}

export function getPPTElementAnimationTimeFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTElementAnimationTime(value)
}

export function getPPTElementAnimationOrderFromJSONValue(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return undefined
  }

  return clampPPTElementAnimationOrder(value)
}
