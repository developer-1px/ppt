import { clamp, type Point } from 'canvas/core'
import { SLIDE_HEIGHT, SLIDE_WIDTH, type Rect } from './retouchModel'

type ClientPointInput = Pick<PointerEvent, 'clientX' | 'clientY'>

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
    slideElement?.querySelectorAll<HTMLElement>('[data-block]') ?? [],
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
