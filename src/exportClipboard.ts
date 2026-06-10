export async function writeExportToClipboard(
  exportCode: string,
  fallbackTextarea: HTMLTextAreaElement | null,
) {
  try {
    await navigator.clipboard.writeText(exportCode)
    return true
  } catch {
    if (fallbackTextarea === null) {
      return false
    }

    fallbackTextarea.value = exportCode
    fallbackTextarea.select()

    return document.execCommand('copy')
  }
}
