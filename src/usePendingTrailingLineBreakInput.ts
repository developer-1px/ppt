import { useRef } from 'react'
import {
  placeCaretAtTextOffset,
  repairTrailingLineBreakInput,
} from './editableTextDom'

export function usePendingTrailingLineBreakInput() {
  const pendingTrailingLineBreakRef = useRef<{ beforeText: string } | null>(null)

  function clearPendingTrailingLineBreak() {
    pendingTrailingLineBreakRef.current = null
  }

  function rememberTrailingLineBreak(lineBreakAtEnd: boolean, beforeText: string) {
    pendingTrailingLineBreakRef.current = lineBreakAtEnd ? { beforeText } : null
  }

  function insertAfterPendingTrailingLineBreak(
    element: HTMLElement,
    insertedText: string,
  ) {
    const pendingLineBreak = pendingTrailingLineBreakRef.current

    pendingTrailingLineBreakRef.current = null

    if (
      !pendingLineBreak ||
      element.textContent !== `${pendingLineBreak.beforeText}\n`
    ) {
      return false
    }

    element.textContent = `${pendingLineBreak.beforeText}\n${insertedText}`
    placeCaretAtTextOffset(
      element,
      pendingLineBreak.beforeText.length + 1 + insertedText.length,
    )

    return true
  }

  function repairPendingTrailingLineBreak(element: HTMLElement) {
    const pendingLineBreak = pendingTrailingLineBreakRef.current

    if (!pendingLineBreak) {
      return
    }

    pendingTrailingLineBreakRef.current = null
    repairTrailingLineBreakInput(element, pendingLineBreak.beforeText)
  }

  return {
    clearPendingTrailingLineBreak,
    insertAfterPendingTrailingLineBreak,
    rememberTrailingLineBreak,
    repairPendingTrailingLineBreak,
  }
}
