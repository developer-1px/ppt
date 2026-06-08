import {
  getDeepTarget,
  isEditableTarget as isKeyboardEditableTarget,
  matchesShortcut,
} from '@interactive-os/keyboard'
import type { Point } from 'canvas/core'
import { GRID_SIZE } from './retouchModel'

type HistoryShortcutAction = 'redo' | 'undo'

const REDO_HISTORY_SHORTCUTS = 'Control+Shift+z Meta+Shift+z Control+y Meta+y'
const UNDO_HISTORY_SHORTCUTS = 'Control+z Meta+z'

export function historyShortcutAction(
  event: KeyboardEvent,
): HistoryShortcutAction | null {
  if (matchesShortcut(event, REDO_HISTORY_SHORTCUTS)) {
    return 'redo'
  }

  if (matchesShortcut(event, UNDO_HISTORY_SHORTCUTS)) {
    return 'undo'
  }

  return null
}

export function isUndoHistoryShortcut(event: KeyboardEvent) {
  return historyShortcutAction(event) === 'undo'
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
