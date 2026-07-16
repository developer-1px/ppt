import {
  createCanvasCssBoundsTransform,
  createCanvasSvgFreehandPathData,
  createCanvasSvgPathData,
  createCanvasSvgPathSegmentData,
  createCanvasSvgBoundsTransform,
  escapeCanvasXmlAttribute,
  formatCanvasSvgNumber,
  type CanvasSvgPathSegment,
} from 'canvas/renderer'

export const createPPTCanvasCssBoundsTransform =
  createCanvasCssBoundsTransform
export const createPPTCanvasSvgBoundsTransform =
  createCanvasSvgBoundsTransform
export const createPPTCanvasSvgFreehandPathData =
  createCanvasSvgFreehandPathData
export const createPPTCanvasSvgPathData = createCanvasSvgPathData
export const createPPTCanvasSvgPathSegmentData =
  createCanvasSvgPathSegmentData
export const escapePPTCanvasXmlAttribute = escapeCanvasXmlAttribute
export const formatPPTCanvasSvgNumber = formatCanvasSvgNumber
export type PPTCanvasSvgPathSegment = CanvasSvgPathSegment
