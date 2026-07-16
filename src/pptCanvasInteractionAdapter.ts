import {
  getNextPPTCanvasDrawingPoints,
  getPPTCanvasPointerStartProjection,
  PPT_LASER_TRAIL_OVERLAY_MODEL,
  previewPPTCanvasPointerLaserInteraction,
  previewPPTCanvasPointerPanInteraction,
  startPPTCanvasPointerLaserInteraction,
  startPPTCanvasPointerPanInteraction,
  usePPTCanvasAppStageElement,
  type PPTCanvasPointerLaserInteractionBase,
  type PPTCanvasPointerPanInteractionBase,
} from './pptCanvasAppAffordanceAdapter'

export { PPT_LASER_TRAIL_OVERLAY_MODEL }
export const usePPTCanvasStageElement = usePPTCanvasAppStageElement
export const getNextPPTDrawingPoints = getNextPPTCanvasDrawingPoints
export const getPPTPointerStartProjection = getPPTCanvasPointerStartProjection
export const previewPPTPointerLaserInteraction =
  previewPPTCanvasPointerLaserInteraction
export const startPPTPointerLaserInteraction =
  startPPTCanvasPointerLaserInteraction
export const previewPPTPointerPanInteraction =
  previewPPTCanvasPointerPanInteraction
export const startPPTPointerPanInteraction =
  startPPTCanvasPointerPanInteraction

export type PPTPointerLaserInteraction = PPTCanvasPointerLaserInteractionBase
export type PPTPointerPanInteraction = PPTCanvasPointerPanInteractionBase
