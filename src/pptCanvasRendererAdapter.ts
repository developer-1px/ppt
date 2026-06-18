import {
  createCanvasCssBoundsTransform,
  createCanvasSvgBoundsTransform,
  escapeCanvasXmlAttribute,
  formatCanvasSvgNumber,
} from 'canvas/renderer'
import {
  createCanvasSvgFreehandPathData,
  createCanvasSvgPathData,
} from 'canvas/renderer/svg-drawing-primitives'

export const createPPTCanvasCssBoundsTransform =
  createCanvasCssBoundsTransform
export const createPPTCanvasSvgBoundsTransform =
  createCanvasSvgBoundsTransform
export const createPPTCanvasSvgFreehandPathData =
  createCanvasSvgFreehandPathData
export const createPPTCanvasSvgPathData = createCanvasSvgPathData
export const escapePPTCanvasXmlAttribute = escapeCanvasXmlAttribute
export const formatPPTCanvasSvgNumber = formatCanvasSvgNumber
