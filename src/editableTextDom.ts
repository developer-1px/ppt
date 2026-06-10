import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  type Rect,
} from './retouchModel'

const CARET_PLACEHOLDER = '\u200B'

type StageScrollPosition = {
  element: HTMLElement
  left: number
  top: number
}

export function applyAutoHeightStyle(
  element: HTMLElement,
  rect: Rect,
  minimumHeight: number,
) {
  element.style.left = `${(rect.x / SLIDE_WIDTH) * 100}%`
  element.style.top = `${(rect.y / SLIDE_HEIGHT) * 100}%`
  element.style.width = `${(rect.width / SLIDE_WIDTH) * 100}%`
  element.style.height = 'auto'
  element.style.minHeight =
    minimumHeight > 0 ? `${(minimumHeight / SLIDE_HEIGHT) * 100}%` : ''
}

export function autoHeightRect(
  element: HTMLElement,
  rect: Rect,
  minimumHeight: number,
): Rect {
  const elementBox = element.getBoundingClientRect()

  if (elementBox.width === 0) {
    return rect
  }

  const slideUnitsPerCssPixel = rect.width / elementBox.width
  const contentHeight = elementBox.height * slideUnitsPerCssPixel
  const height = Math.max(minimumHeight, Math.ceil(contentHeight))
  const y = Math.min(rect.y, Math.max(0, SLIDE_HEIGHT - height))

  return rect.height === height && rect.y === y ? rect : { ...rect, y, height }
}

export function placeCaretAtEnd(root: HTMLElement) {
  const selection = window.getSelection()
  const range = document.createRange()
  range.selectNodeContents(root)
  range.collapse(false)
  selection?.removeAllRanges()
  selection?.addRange(range)
}

export function isSelectionAtTextEnd(root: HTMLElement) {
  const selection = window.getSelection()

  if (selection === null || selection.rangeCount === 0) {
    return true
  }

  const range = selection.getRangeAt(0)

  if (!root.contains(range.commonAncestorContainer)) {
    return true
  }

  const afterCaret = range.cloneRange()
  afterCaret.selectNodeContents(root)
  afterCaret.setStart(range.endContainer, range.endOffset)

  return afterCaret.toString().length === 0
}

export function insertTextAtSelection(root: HTMLElement, text: string) {
  const selection = window.getSelection()
  const insertedText = text

  if (selection === null || selection.rangeCount === 0) {
    root.append(document.createTextNode(insertedText))
    placeCaretAtEnd(root)
    return
  }

  const range = selection.getRangeAt(0)

  if (!root.contains(range.commonAncestorContainer)) {
    root.append(document.createTextNode(insertedText))
    placeCaretAtEnd(root)
    return
  }

  const textNode = document.createTextNode(insertedText)

  range.deleteContents()
  range.insertNode(textNode)
  range.setStart(textNode, insertedText.length)
  range.collapse(true)
  selection.removeAllRanges()
  selection.addRange(range)
}

export function repairTrailingLineBreakInput(root: HTMLElement, beforeText: string) {
  const text = root.textContent ?? ''

  if (text === `${beforeText}\n`) {
    return
  }

  if (!text.startsWith(beforeText) || !text.endsWith('\n')) {
    return
  }

  const insertedText = text.slice(beforeText.length, -1)

  if (insertedText.length === 0) {
    return
  }

  root.textContent = `${beforeText}\n${insertedText}`
  placeCaretAtTextOffset(root, beforeText.length + 1 + insertedText.length)
}

export function placeCaretAtTextOffset(root: HTMLElement, offset: number) {
  const selection = window.getSelection()
  const range = document.createRange()
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
  let remainingOffset = offset
  let node = walker.nextNode()

  while (node) {
    const textLength = node.textContent?.length ?? 0

    if (remainingOffset <= textLength) {
      range.setStart(node, remainingOffset)
      range.collapse(true)
      selection?.removeAllRanges()
      selection?.addRange(range)
      return
    }

    remainingOffset -= textLength
    node = walker.nextNode()
  }

  placeCaretAtEnd(root)
}

export function normalizeEditableText(text: string) {
  return text.replaceAll(CARET_PLACEHOLDER, '')
}

export function placeCaretFromPoint(root: HTMLElement, x: number, y: number) {
  const range = readCaretRangeFromPoint(x, y)

  if (range === null || !root.contains(range.startContainer)) {
    return false
  }

  const selection = window.getSelection()
  selection?.removeAllRanges()
  selection?.addRange(range)

  return true
}

function readCaretRangeFromPoint(x: number, y: number) {
  if ('caretPositionFromPoint' in document) {
    const position = document.caretPositionFromPoint(x, y)

    if (position !== null) {
      const range = document.createRange()
      range.setStart(position.offsetNode, position.offset)
      range.collapse(true)

      return range
    }
  }

  if ('caretRangeFromPoint' in document) {
    return document.caretRangeFromPoint(x, y)
  }

  return null
}

export function rememberStageScroll(element: HTMLElement): StageScrollPosition | null {
  const stage = element.closest<HTMLElement>('.stage-shell')

  return stage
    ? {
        element: stage,
        left: stage.scrollLeft,
        top: stage.scrollTop,
      }
    : null
}

export function restoreStageScroll(position: StageScrollPosition | null) {
  if (position === null) {
    return
  }

  position.element.scrollLeft = position.left
  position.element.scrollTop = position.top
}
