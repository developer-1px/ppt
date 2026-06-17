import { useCanvasAppStageElement } from 'canvas/app/stage-element'
import { getNextCanvasDrawingPoints } from 'canvas/app/pointer-drawing'
import {
  CANVAS_LASER_TRAIL_OVERLAY_MODEL,
  previewCanvasPointerLaserInteraction,
  startCanvasPointerLaserInteraction,
  type CanvasPointerLaserInteraction,
} from 'canvas/app/pointer-laser'
import {
  previewCanvasPointerPanInteraction,
  startCanvasPointerPanInteraction,
  type CanvasPointerPanInteraction,
} from 'canvas/app/pointer-pan-interaction'
import { getCanvasPointerStartProjection } from 'canvas/app/pointer-start-session'

export const usePPTCanvasStageElement = useCanvasAppStageElement
export const getNextPPTDrawingPoints = getNextCanvasDrawingPoints
export const PPT_LASER_TRAIL_OVERLAY_MODEL = CANVAS_LASER_TRAIL_OVERLAY_MODEL
export const getPPTPointerStartProjection = getCanvasPointerStartProjection
export const previewPPTPointerLaserInteraction =
  previewCanvasPointerLaserInteraction
export const startPPTPointerLaserInteraction =
  startCanvasPointerLaserInteraction
export const previewPPTPointerPanInteraction =
  previewCanvasPointerPanInteraction
export const startPPTPointerPanInteraction = startCanvasPointerPanInteraction

export type PPTPointerLaserInteraction = CanvasPointerLaserInteraction
export type PPTPointerPanInteraction = CanvasPointerPanInteraction
