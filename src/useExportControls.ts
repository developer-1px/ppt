import { useEffect, useRef, useState } from 'react'
import { writeExportToClipboard } from './exportClipboard'

export function useExportControls({
  exportCode,
  readCommittedExportCode,
  statusMatchesVisibleSlide,
}: {
  exportCode: string
  readCommittedExportCode: () => string
  statusMatchesVisibleSlide: boolean
}) {
  const exportTextareaRef = useRef<HTMLTextAreaElement | null>(null)
  const previousExportCodeRef = useRef<string | null>(null)
  const [copiedExportCode, setCopiedExportCode] = useState<string | null>(null)
  const [downloadedExportCode, setDownloadedExportCode] = useState<string | null>(null)
  const [failedCopyExportCode, setFailedCopyExportCode] = useState<string | null>(null)

  const exportCopied = statusMatchesVisibleSlide && copiedExportCode === exportCode
  const exportDownloaded =
    statusMatchesVisibleSlide && downloadedExportCode === exportCode
  const exportCopyFailed =
    statusMatchesVisibleSlide && failedCopyExportCode === exportCode
  const copyState: 'copied' | 'failed' | 'idle' = exportCopied
    ? 'copied'
    : exportCopyFailed
      ? 'failed'
      : 'idle'
  const copyTitle =
    copyState === 'copied'
      ? 'Copied'
      : copyState === 'failed'
        ? 'Copy failed'
        : 'Copy HTML'

  useEffect(() => {
    if (previousExportCodeRef.current === null) {
      previousExportCodeRef.current = exportCode
      return
    }

    if (previousExportCodeRef.current === exportCode) {
      return
    }

    previousExportCodeRef.current = exportCode
    setFailedCopyExportCode(null)
  }, [exportCode])

  async function copyExportCode() {
    const nextExportCode = readCommittedExportCode()
    const copied = await writeExportToClipboard(nextExportCode, exportTextareaRef.current)

    if (!copied) {
      setCopiedExportCode(null)
      setFailedCopyExportCode(nextExportCode)
      return
    }

    setFailedCopyExportCode(null)
    setCopiedExportCode(nextExportCode)
  }

  function downloadExportCode() {
    const nextExportCode = readCommittedExportCode()
    const url = URL.createObjectURL(
      new Blob([nextExportCode], { type: 'text/html;charset=utf-8' }),
    )
    const anchor = document.createElement('a')

    anchor.href = url
    anchor.download = 'retouched-slides.html'
    anchor.rel = 'noopener'
    document.body.append(anchor)
    anchor.click()
    anchor.remove()
    URL.revokeObjectURL(url)
    setFailedCopyExportCode(null)
    setDownloadedExportCode(nextExportCode)
  }

  function resetExportFeedback() {
    setCopiedExportCode(null)
    setDownloadedExportCode(null)
    setFailedCopyExportCode(null)
  }

  return {
    copyExportCode,
    copyState,
    copyTitle,
    downloadExportCode,
    exportCopied,
    exportDownloaded,
    exportTextareaRef,
    resetExportFeedback,
  }
}
