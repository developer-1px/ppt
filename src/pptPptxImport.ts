import JSZip from 'jszip'
import { PPTDeckSchema, type PPTDeck } from './pptModel'
import {
  PPTX_MIME_TYPE,
  PPTX_MODEL_CUSTOM_XML_CONTENT_TYPE,
  PPTX_MODEL_CUSTOM_XML_NAMESPACE,
  PPTX_MODEL_CUSTOM_XML_PATH,
} from './pptPptxExport'

export const PPTX_DECK_MODEL_IMPORT_FORMAT = 'pptx-custom-xml-ppt-deck' as const

export type PPTDeckPPTXImportResult = {
  deck: PPTDeck
  format: typeof PPTX_DECK_MODEL_IMPORT_FORMAT
  jsonLength: number
}

export async function importPPTDeckFromPPTXBlob(
  blob: Blob,
): Promise<PPTDeckPPTXImportResult | null> {
  try {
    const zip = await JSZip.loadAsync(await blob.arrayBuffer())
    const xml = await zip.file(PPTX_MODEL_CUSTOM_XML_PATH)?.async('string')

    if (!xml) {
      return null
    }

    const payload = getPPTDeckModelPayloadFromCustomXml(xml)

    if (!payload) {
      return null
    }

    const parsed = PPTDeckSchema.safeParse(JSON.parse(payload))

    if (!parsed.success) {
      return null
    }

    return {
      deck: parsed.data,
      format: PPTX_DECK_MODEL_IMPORT_FORMAT,
      jsonLength: payload.length,
    }
  } catch {
    return null
  }
}

export function getPPTDeckPPTXFileFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTDeckPPTXFilesFromDataTransfer(dataTransfer)[0] ?? null
}

export function canImportPPTDeckPPTXFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getPPTDeckPPTXFilesFromDataTransfer(dataTransfer).length > 0 ||
    hasPPTDeckPPTXItem(dataTransfer)
}

function getPPTDeckPPTXFilesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return Array.from(dataTransfer?.files ?? [])
    .filter(isPPTDeckPPTXFile)
}

function hasPPTDeckPPTXItem(dataTransfer: DataTransfer | null) {
  return Array.from(dataTransfer?.items ?? [])
    .some((item) => item.kind === 'file' && item.type === PPTX_MIME_TYPE)
}

function isPPTDeckPPTXFile(file: File) {
  return file.type === PPTX_MIME_TYPE ||
    file.name.toLowerCase().endsWith('.pptx')
}

function getPPTDeckModelPayloadFromCustomXml(xml: string) {
  const match = xml.match(/<pptDeck\b([^>]*)>([\s\S]*?)<\/pptDeck>/)

  if (!match) {
    return null
  }

  const attributes = match[1]

  if (!attributes.includes(`xmlns="${PPTX_MODEL_CUSTOM_XML_NAMESPACE}"`) ||
    !attributes.includes(`contentType="${PPTX_MODEL_CUSTOM_XML_CONTENT_TYPE}"`)) {
    return null
  }

  return unescapePPTXXmlText(match[2]).trim()
}

function unescapePPTXXmlText(value: string) {
  return value
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}
