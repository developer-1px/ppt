import type { JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import {
  RetouchDeckSchema,
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  blockPointer,
  blockTextPointer,
  slideBlocksPointer,
  slidePointer,
  type RetouchDeck,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'

const RETOUCH_SURFACE_CONTRACT =
  'interactive-os.retouch-surface.v1'

const RETOUCH_SURFACE_POINTER_TEMPLATES = {
  root: '',
  slides: '/slides',
  slide: '/slides/{slideIndex}',
  blocks: '/slides/{slideIndex}/blocks',
  block: '/slides/{slideIndex}/blocks/{blockIndex}',
  blockIdentity: '/id',
  blockText: '/text',
} as const

const RETOUCH_SURFACE_CAPABILITIES = [
  'text-edit',
  'autoheight',
  'object-selection',
  'marquee',
  'move',
  'resize',
  'z-order',
  'export-patch-manifest',
] as const

export type RetouchSurfaceCapability =
  (typeof RETOUCH_SURFACE_CAPABILITIES)[number]

export type RetouchSurfaceDocument = RetouchDeck
export type RetouchSurfaceSlide = RetouchSlide
export type RetouchSurfaceBlock = SlideBlock

export type RetouchSurfaceBlockIdentity = {
  field: 'id'
  path: typeof RETOUCH_SURFACE_POINTER_TEMPLATES.blockIdentity
}

export type RetouchSurfacePaths = {
  root: Pointer
  slides: Pointer
  slide: (slideIndex: number) => Pointer
  blocks: (slideIndex: number) => Pointer
  block: (slideIndex: number, blockIndex: number) => Pointer
  blockText: (blockPointerValue: Pointer) => Pointer
  templates: typeof RETOUCH_SURFACE_POINTER_TEMPLATES
}

export type RetouchSurfaceDescriptor = {
  blockIdentity: RetouchSurfaceBlockIdentity
  capabilities: readonly RetouchSurfaceCapability[]
  contract: typeof RETOUCH_SURFACE_CONTRACT
  paths: RetouchSurfacePaths
  schema: typeof RetouchDeckSchema
  slideSize: {
    height: typeof SLIDE_HEIGHT
    width: typeof SLIDE_WIDTH
  }
}

export type RetouchSurfaceChangeKind =
  | 'arrange'
  | 'import'
  | 'layer-order'
  | 'layout'
  | 'reset'
  | 'selection'
  | 'text'

export type RetouchSurfaceRuntimeChange = {
  focusPointer?: Pointer | null
  kind: RetouchSurfaceChangeKind
  label: string
  patch: readonly JSONPatchOperation[]
  selection?: SelectionSnap
}

export type RetouchPatchCommitOptions =
  | string
  | { label: string; mergeKey?: string; selection?: SelectionSnap }

export type RetouchPatchCommit = (
  patch: readonly JSONPatchOperation[],
  options: RetouchPatchCommitOptions,
) => void

export type RetouchSurfaceCommitPatch = (
  patch: JSONPatchOperation[],
  pointer: Pointer,
  label: string,
  mergeKey?: string,
  selection?: SelectionSnap,
) => void

export const PPT_RETOUCH_SURFACE: RetouchSurfaceDescriptor = {
  blockIdentity: {
    field: 'id',
    path: RETOUCH_SURFACE_POINTER_TEMPLATES.blockIdentity,
  },
  capabilities: RETOUCH_SURFACE_CAPABILITIES,
  contract: RETOUCH_SURFACE_CONTRACT,
  paths: {
    root: RETOUCH_SURFACE_POINTER_TEMPLATES.root,
    slides: RETOUCH_SURFACE_POINTER_TEMPLATES.slides,
    slide: slidePointer,
    blocks: slideBlocksPointer,
    block: blockPointer,
    blockText: blockTextPointer,
    templates: RETOUCH_SURFACE_POINTER_TEMPLATES,
  },
  schema: RetouchDeckSchema,
  slideSize: {
    height: SLIDE_HEIGHT,
    width: SLIDE_WIDTH,
  },
}
