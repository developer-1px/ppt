import {
  getDeepTarget,
  isEditableTarget as isKeyboardEditableTarget,
  matchesShortcut,
} from '@interactive-os/keyboard'
import { GRID_SIZE } from './retouchModel'
import type { Point } from './layoutInteraction'

export function isHistoryShortcut(event: KeyboardEvent) {
  return matchesShortcut(
    event,
    'Control+z Meta+z Control+Shift+z Meta+Shift+z Control+y Meta+y',
  )
}

export function isEditableTarget(target: EventTarget | Event | null) {
  if (isKeyboardEditableTarget(target)) {
    return true
  }

  const deepTarget = target instanceof Event ? getDeepTarget(target) : target

  if (!(deepTarget instanceof HTMLElement)) {
    return false
  }

  return Boolean(deepTarget.closest('select'))
}

export function isControlTarget(target: EventTarget | Event | null) {
  const deepTarget = target instanceof Event ? getDeepTarget(target) : target

  if (!(deepTarget instanceof HTMLElement)) {
    return false
  }

  return Boolean(deepTarget.closest('button, [role="tab"], [role="toolbar"]'))
}

export function arrowKeyDelta(key: string, largeStep: boolean): Point | null {
  const step = largeStep ? GRID_SIZE * 5 : GRID_SIZE

  if (key === 'ArrowLeft') {
    return { x: -step, y: 0 }
  }

  if (key === 'ArrowRight') {
    return { x: step, y: 0 }
  }

  if (key === 'ArrowUp') {
    return { x: 0, y: -step }
  }

  if (key === 'ArrowDown') {
    return { x: 0, y: step }
  }

  return null
}
