import type { Point } from 'canvas/core'
import type { Pointer } from 'zod-crud'

export type CanvasView = 'slide' | 'grid'
export type RetouchMode = 'text' | 'layout'
export type ResetScope = 'deck' | 'layout' | 'text'

export type EditingState = {
  clientPoint?: Point
  pointer: Pointer
}
