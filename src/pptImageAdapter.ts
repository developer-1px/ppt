import {
  createSlideEditObjectImageCropDescriptor,
  createSlideEditObjectImageReplaceDescriptor,
  normalizeSlideEditObjectImageCropFit,
  type SlideEditObjectImageCropDescriptor,
  type SlideEditObjectImageReplaceDescriptor,
} from './pptSlideEditAffordanceAdapter'
import type { PPTImage, PPTImageCrop, PPTImageFit } from './pptModel'

export function getPPTImageFit(element: PPTImage): PPTImageFit {
  return element.fit ?? 'cover'
}

export function normalizePPTImageFit(
  value: string | null | undefined,
): PPTImageFit {
  return normalizeSlideEditObjectImageCropFit(value)
}

export function isPPTImageFit(value: string): value is PPTImageFit {
  return value === 'cover' || value === 'contain'
}

export function getPPTImageCrop(element: PPTImage): PPTImageCrop {
  return element.crop ?? { x: 50, y: 50 }
}

export function getPPTImageCropDescriptor(
  slideId: string,
  element: PPTImage,
): SlideEditObjectImageCropDescriptor<string, string> {
  return createSlideEditObjectImageCropDescriptor({
    crop: getPPTImageCrop(element),
    fit: getPPTImageFit(element),
    objectId: element.id,
    slideId,
  })
}

export function getPPTImageReplaceDescriptor(
  slideId: string,
  element: PPTImage,
): SlideEditObjectImageReplaceDescriptor<string, string> {
  return createSlideEditObjectImageReplaceDescriptor({
    isSupported: element.locked !== true && element.visible !== false,
    objectId: element.id,
    slideId,
    sourceName: element.name,
    unsupportedReason: element.locked === true
      ? 'locked-object'
      : 'unsupported-object',
  })
}
