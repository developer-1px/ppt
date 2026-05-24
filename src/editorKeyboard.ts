import { GRID_SIZE } from './retouchModel'
import type { Point } from './layoutInteraction'

export function isHistoryShortcut(event: Pick<KeyboardEvent, 'altKey' | 'ctrlKey' | 'metaKey'>) {
  return (event.metaKey || event.ctrlKey) && !event.altKey
}

export function isEditableTarget(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  return Boolean(target.closest('[contenteditable], textarea, input, select'))
}

export function isControlTarget(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  return Boolean(target.closest('button, [role="tab"], [role="toolbar"]'))
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
