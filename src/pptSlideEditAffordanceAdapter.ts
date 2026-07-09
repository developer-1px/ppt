export * from '@interactive-os/slide-edit-affordance'
import type {
  SlideEditClipboardOperation,
} from '@interactive-os/slide-edit-affordance'

export const SLIDE_EDIT_SLIDE_CLIPBOARD_MIME_TYPE =
  'application/vnd.interactive-os.slide-edit.slide+json'
export const SLIDE_EDIT_SLIDE_CLIPBOARD_HTML_SCRIPT_ATTRIBUTE =
  'data-slide-edit-slide-clipboard-json'

export type SlideEditSlideClipboardOperation = SlideEditClipboardOperation

export type SlideEditSlideClipboardMetadata<
  TSlideId extends string = string,
> = {
  objectCount: number
  slideId: TSlideId
  title: string
}

export type SlideEditSlideClipboardPayload<
  TSlideId extends string = string,
  TSlide = unknown,
> = {
  activeSlideId: TSlideId
  metadata: readonly SlideEditSlideClipboardMetadata<TSlideId>[]
  operation: SlideEditSlideClipboardOperation
  selectedSlideIds: readonly TSlideId[]
  slides: readonly TSlide[]
  sourceSlideId: TSlideId
  type: 'slide-clipboard'
}

export type SlideEditSlideClipboardPasteTarget<TSlideId extends string = string> =
  | {
    kind: 'after-slide'
    slideId: TSlideId | null
  }
  | {
    kind: 'end'
  }
  | {
    kind: 'index'
    index: number
  }

export type SlideEditSlideClipboardPastePlacement<
  TSlideId extends string = string,
> = {
  afterSlideId: TSlideId | null
  insertIndex: number
  target: SlideEditSlideClipboardPasteTarget<TSlideId>
}

export type SlideEditSlideClipboardObjectMapping<
  TObjectId extends string = string,
> = {
  objectIndex: number
  sourceObjectId: TObjectId
  targetObjectId: TObjectId
}

export type SlideEditSlideClipboardPasteMapping<
  TSlideId extends string = string,
  TObjectId extends string = string,
> = {
  objectMappings: readonly SlideEditSlideClipboardObjectMapping<TObjectId>[]
  slideIndex: number
  sourceSlideId: TSlideId
  targetSlideId: TSlideId
}

export type SlideEditSlideClipboardPastePlan<
  TSlideId extends string = string,
  TObjectId extends string = string,
> = {
  afterSlideId: TSlideId | null
  insertIndex: number
  mappings: readonly SlideEditSlideClipboardPasteMapping<TSlideId, TObjectId>[]
  operation: SlideEditSlideClipboardOperation
  sourceSlideId: TSlideId
  targetSlideIds: readonly TSlideId[]
}

export type SlideEditSlideClipboardPasteCommand<
  TSlideId extends string = string,
  TObjectId extends string = string,
  TSlide = unknown,
> = {
  id: 'paste-slides'
  pastePlan: SlideEditSlideClipboardPastePlan<TSlideId, TObjectId>
  payload: SlideEditSlideClipboardPayload<TSlideId, TSlide>
}

export type SlideEditSlideClipboardPasteHostCommandEffect<
  TSlideId extends string = string,
  TObjectId extends string = string,
  TSlide = unknown,
> = {
  payload: SlideEditSlideClipboardPasteCommand<TSlideId, TObjectId, TSlide>
  selection: {
    slideId: TSlideId
    slideIds: readonly TSlideId[]
  }
  type: 'slide-command-effect'
}

export type SlideEditMarkdownListItem = {
  level?: number
  text: string
}

export type SlideEditMarkdownSlideBlock =
  | {
    kind: 'paragraph'
    markdown: string
    text: string
  }
  | {
    items: readonly SlideEditMarkdownListItem[]
    kind: 'ordered-list' | 'unordered-list'
  }

export type SlideEditMarkdownSlideSource = {
  body: readonly SlideEditMarkdownSlideBlock[]
  notes: readonly SlideEditMarkdownSlideBlock[]
  title: string
}

export type SlideEditMarkdownDeckSource = {
  deckTitle?: string | null
  payloadLength: number
  slides: readonly SlideEditMarkdownSlideSource[]
}

export function createSlideEditSlideClipboardPayload<
  TSlideId extends string,
  TSlide,
>({
  activeSlideId,
  getObjectCount,
  getSlideId,
  getTitle,
  operation = 'copy',
  selectedSlideIds,
  slides,
  sourceSlideId,
}: {
  activeSlideId: TSlideId
  getObjectCount: (slide: TSlide) => number
  getSlideId: (slide: TSlide) => TSlideId
  getTitle: (slide: TSlide) => string
  operation?: SlideEditSlideClipboardOperation
  selectedSlideIds: readonly TSlideId[]
  slides: readonly TSlide[]
  sourceSlideId: TSlideId
}): SlideEditSlideClipboardPayload<TSlideId, TSlide> | null {
  if (slides.length === 0 || selectedSlideIds.length === 0) {
    return null
  }

  return {
    activeSlideId,
    metadata: slides.map((slide) => ({
      objectCount: getObjectCount(slide),
      slideId: getSlideId(slide),
      title: getTitle(slide),
    })),
    operation,
    selectedSlideIds,
    slides,
    sourceSlideId,
    type: 'slide-clipboard',
  }
}

export function parseSlideEditSlideClipboardPayload<TSlide = unknown>(
  value: unknown,
): SlideEditSlideClipboardPayload<string, TSlide> | null {
  if (!isSlideEditRecord(value)) {
    return null
  }

  const payloadValue = value.type === 'slide-clipboard'
    ? value
    : isSlideEditRecord(value.payload) &&
      value.payload.type === 'slide-clipboard'
      ? value.payload
      : null

  if (!isSlideEditRecord(payloadValue) || !Array.isArray(payloadValue.slides)) {
    return null
  }

  const metadata = Array.isArray(payloadValue.metadata)
    ? payloadValue.metadata.flatMap((item): SlideEditSlideClipboardMetadata[] => {
        if (!isSlideEditRecord(item) || typeof item.slideId !== 'string') {
          return []
        }

        return [{
          objectCount: typeof item.objectCount === 'number'
            ? item.objectCount
            : 0,
          slideId: item.slideId,
          title: typeof item.title === 'string' ? item.title : item.slideId,
        }]
      })
    : []
  const selectedSlideIds = normalizeSlideEditStringArray(
    payloadValue.selectedSlideIds,
  )
  const sourceSlideId = typeof payloadValue.sourceSlideId === 'string'
    ? payloadValue.sourceSlideId
    : metadata[0]?.slideId ?? selectedSlideIds[0] ?? ''
  const activeSlideId = typeof payloadValue.activeSlideId === 'string'
    ? payloadValue.activeSlideId
    : selectedSlideIds[0] ?? sourceSlideId

  return {
    activeSlideId,
    metadata,
    operation: payloadValue.operation === 'cut' ? 'cut' : 'copy',
    selectedSlideIds,
    slides: payloadValue.slides as TSlide[],
    sourceSlideId,
    type: 'slide-clipboard',
  }
}

export function resolveSlideEditSlideClipboardPastePlacement<
  TSlideId extends string,
>({
  slideOrder,
  target,
}: {
  slideOrder: readonly TSlideId[]
  target: SlideEditSlideClipboardPasteTarget<TSlideId>
}): SlideEditSlideClipboardPastePlacement<TSlideId> {
  if (target.kind === 'index') {
    const insertIndex = clampSlideEditSlideClipboardInsertIndex(
      target.index,
      slideOrder.length,
    )

    return {
      afterSlideId: slideOrder[insertIndex - 1] ?? null,
      insertIndex,
      target,
    }
  }

  if (target.kind === 'end') {
    return {
      afterSlideId: slideOrder.at(-1) ?? null,
      insertIndex: slideOrder.length,
      target,
    }
  }

  const targetIndex = target.slideId
    ? slideOrder.indexOf(target.slideId)
    : -1
  const insertIndex = targetIndex >= 0 ? targetIndex + 1 : slideOrder.length

  return {
    afterSlideId: targetIndex >= 0 ? target.slideId : slideOrder.at(-1) ?? null,
    insertIndex,
    target,
  }
}

export function createSlideEditSlideClipboardPasteCommandEffect<
  TSlideId extends string,
  TObjectId extends string,
  TSlide,
>({
  payload,
  placement,
  remapPolicy,
}: {
  payload: SlideEditSlideClipboardPayload<TSlideId, TSlide>
  placement: SlideEditSlideClipboardPastePlacement<TSlideId>
  remapPolicy: {
    createObjectId: (input: {
      objectIndex: number
      slideIndex: number
      sourceObjectId: TObjectId
      sourceSlide: TSlide
      sourceSlideId: TSlideId
    }) => TObjectId
    createSlideId: (input: {
      slideIndex: number
      sourceSlide: TSlide
      sourceSlideId: TSlideId
    }) => TSlideId
    getObjectIds: (slide: TSlide) => readonly TObjectId[]
  }
}): SlideEditSlideClipboardPasteHostCommandEffect<TSlideId, TObjectId, TSlide> | null {
  const mappings = payload.slides.map((slide, slideIndex) => {
    const sourceSlideId = payload.metadata[slideIndex]?.slideId ??
      payload.selectedSlideIds[slideIndex] ??
      payload.sourceSlideId
    const targetSlideId = remapPolicy.createSlideId({
      slideIndex,
      sourceSlide: slide,
      sourceSlideId,
    })

    return {
      objectMappings: remapPolicy.getObjectIds(slide).map((
        sourceObjectId,
        objectIndex,
      ) => ({
        objectIndex,
        sourceObjectId,
        targetObjectId: remapPolicy.createObjectId({
          objectIndex,
          slideIndex,
          sourceObjectId,
          sourceSlide: slide,
          sourceSlideId,
        }),
      })),
      slideIndex,
      sourceSlideId,
      targetSlideId,
    }
  })
  const targetSlideIds = mappings.map((mapping) => mapping.targetSlideId)
  const selectedSlideId = targetSlideIds[0]

  if (!selectedSlideId) {
    return null
  }

  return {
    payload: {
      id: 'paste-slides',
      pastePlan: {
        afterSlideId: placement.afterSlideId,
        insertIndex: placement.insertIndex,
        mappings,
        operation: payload.operation,
        sourceSlideId: payload.sourceSlideId,
        targetSlideIds,
      },
      payload,
    },
    selection: {
      slideId: selectedSlideId,
      slideIds: targetSlideIds,
    },
    type: 'slide-command-effect',
  }
}

export function mapSlideEditSlideClipboardPasteSlides<
  TSlideId extends string,
  TObjectId extends string,
  TSlide,
  TResult,
>({
  pastePlan,
  payload,
  transform,
}: {
  pastePlan: SlideEditSlideClipboardPastePlan<TSlideId, TObjectId>
  payload: Pick<SlideEditSlideClipboardPayload<TSlideId, TSlide>, 'slides'>
  transform: (input: {
    mapping: SlideEditSlideClipboardPasteMapping<TSlideId, TObjectId>
    slideIndex: number
    source: TSlide
  }) => TResult | null | undefined
}): TResult[] {
  const results: TResult[] = []

  pastePlan.mappings.forEach((mapping, slideIndex) => {
    const source = payload.slides[slideIndex]

    if (source === undefined) {
      return
    }

    const result = transform({ mapping, slideIndex, source })

    if (result !== null && result !== undefined) {
      results.push(result)
    }
  })

  return results
}

export function getSlideEditMarkdownDeckSource(input?: {
  markdown: string
  sourceType?: string
}): SlideEditMarkdownDeckSource | null {
  void input

  return null
}

export function getSlideEditMarkdownDeckSourceFromDataTransfer(input?: {
  dataTransfer: DataTransfer | null
}): SlideEditMarkdownDeckSource | null {
  void input

  return null
}

function clampSlideEditSlideClipboardInsertIndex(
  value: number,
  slideCount: number,
) {
  if (!Number.isFinite(value)) {
    return slideCount
  }

  return Math.max(0, Math.min(slideCount, Math.round(value)))
}

function isSlideEditRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function normalizeSlideEditStringArray(value: unknown): string[] {
  return Array.isArray(value)
    ? value.filter((item): item is string => typeof item === 'string')
    : []
}
