import { clamp, type Point } from 'canvas/core'
import { HTML_SLIDE_ATTRIBUTES } from './htmlSlideContract'
import { PLAIN_TEXT_BLOCK_EDITOR_SELECTOR } from './plainTextBlockEditor'
import { SLIDE_HEIGHT, SLIDE_WIDTH, type Rect } from './retouchModel'

type ClientPointInput = Pick<PointerEvent, 'clientX' | 'clientY'>

type PlainTextBlockEditorElements = {
  blockElement: HTMLElement
  editorElement: HTMLElement
}

const SLIDE_BLOCK_SELECTOR = `[${HTML_SLIDE_ATTRIBUTES.block}]`
const RETOUCH_SELECTION_CONTROL_SELECTOR = '.selection-overlay, .resize-handle'
const RETOUCH_MARQUEE_SELECTION_SELECTOR = '.marquee-selection'
const RETOUCH_MARQUEE_START_IGNORE_SELECTOR = `${SLIDE_BLOCK_SELECTOR}, ${RETOUCH_SELECTION_CONTROL_SELECTOR}`
const RETOUCH_STAGE_BACKGROUND_IGNORE_SELECTOR = `${RETOUCH_MARQUEE_START_IGNORE_SELECTOR}, ${RETOUCH_MARQUEE_SELECTION_SELECTOR}`

export function closestSlideBlockElement(
  target: EventTarget | null,
): HTMLElement | null {
  const element = closestElement(target, SLIDE_BLOCK_SELECTOR)

  return element instanceof HTMLElement ? element : null
}

export function isRetouchMarqueeStartIgnoredTarget(
  target: EventTarget | null,
) {
  return Boolean(closestElement(target, RETOUCH_MARQUEE_START_IGNORE_SELECTOR))
}

export function isRetouchStageBackgroundIgnoredTarget(
  target: EventTarget | null,
) {
  return Boolean(closestElement(target, RETOUCH_STAGE_BACKGROUND_IGNORE_SELECTOR))
}

export function readPlainTextBlockEditorElements(
  slideElement: HTMLElement | null,
): PlainTextBlockEditorElements | null {
  const editorElement = slideElement?.querySelector<HTMLElement>(
    PLAIN_TEXT_BLOCK_EDITOR_SELECTOR,
  )
  const blockElement = closestSlideBlockElement(editorElement ?? null)

  if (!editorElement || !blockElement) {
    return null
  }

  return {
    blockElement,
    editorElement,
  }
}

export function readSlidePoint(
  slideElement: HTMLElement | null,
  event: ClientPointInput,
): Point | null {
  const rect = slideElement?.getBoundingClientRect()

  if (!rect || rect.width === 0 || rect.height === 0) {
    return null
  }

  return {
    x: clamp(
      ((event.clientX - rect.left) / rect.width) * SLIDE_WIDTH,
      0,
      SLIDE_WIDTH,
    ),
    y: clamp(
      ((event.clientY - rect.top) / rect.height) * SLIDE_HEIGHT,
      0,
      SLIDE_HEIGHT,
    ),
  }
}

export function readSlideBlockRect(
  slideElement: HTMLElement | null,
  blockId: string,
): Rect | null {
  const slideRect = slideElement?.getBoundingClientRect()
  const blockElement = Array.from(
    slideElement?.querySelectorAll<HTMLElement>(SLIDE_BLOCK_SELECTOR) ?? [],
  ).find((element) => element.dataset.block === blockId)
  const blockRect = blockElement?.getBoundingClientRect()

  if (!slideRect || !blockRect || slideRect.width === 0 || slideRect.height === 0) {
    return null
  }

  return {
    x: ((blockRect.left - slideRect.left) / slideRect.width) * SLIDE_WIDTH,
    y: ((blockRect.top - slideRect.top) / slideRect.height) * SLIDE_HEIGHT,
    width: (blockRect.width / slideRect.width) * SLIDE_WIDTH,
    height: (blockRect.height / slideRect.height) * SLIDE_HEIGHT,
  }
}

function closestElement(target: EventTarget | null, selector: string) {
  return target instanceof Element ? target.closest(selector) : null
}
