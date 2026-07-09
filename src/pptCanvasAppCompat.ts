import {
  useMemo,
  useRef,
} from 'react'
import {
  writeCanvasRichClipboardPayload as writeCanvasRichClipboardPayloadBase,
  type CanvasRichClipboardWriteInput as CanvasRichClipboardWriteInputBase,
  type CanvasRichClipboardWriteMode,
} from 'canvas/app-public'
import {
  getCanvasWheelViewport,
} from 'canvas/engine'
import type {
  Bounds,
  Point,
  Viewport,
} from 'canvas/core'

export * from 'canvas/app-public'

export const CANVAS_KEYBOARD_SELECTION_CYCLE_INTENT_MODEL =
  'canvas-keyboard-selection-cycle-intent'
export const CANVAS_KEYBOARD_SELECTION_CYCLE_KEYS = ['Tab'] as const
export const CANVAS_KEYBOARD_SELECTION_CYCLE_MODEL =
  'canvas-keyboard-selection-cycle'
export const CANVAS_KEYBOARD_TEXT_EDIT_START_INTENT_MODEL =
  'canvas-keyboard-text-edit-start-intent'
export const CANVAS_KEYBOARD_TEXT_EDIT_START_KEYS = 'printable'
export const CANVAS_KEYBOARD_TEXT_EDIT_START_MODEL =
  'canvas-keyboard-text-edit-start'
export const CANVAS_KEYBOARD_TEXT_FONT_SIZE_INTENT_MODEL =
  'canvas-keyboard-text-font-size-intent'
export const CANVAS_KEYBOARD_TEXT_FONT_SIZE_KEYS = [
  '>',
  '<',
] as const
export const CANVAS_KEYBOARD_TEXT_FONT_SIZE_MODEL =
  'canvas-keyboard-text-font-size'
export const CANVAS_KEYBOARD_TEXT_FONT_SIZE_STEP = 2
export const CANVAS_IMAGE_FILE_IMPORT_SUPPORTED_FORMATS = [
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/webp',
  'image/svg+xml',
] as const
export const CANVAS_IMAGE_SOURCE_IMPORT_SUPPORTED_FORMATS = [
  'image/svg+xml',
  'text/html',
  'text/plain',
] as const
export const CANVAS_LASER_TRAIL_OVERLAY_MODEL = 'canvas-laser-trail-overlay'
export const CANVAS_RICH_TEXT_PASTE_SUPPORTED_FORMATS = [
  'text/html',
  'text/markdown',
] as const
export const CANVAS_TABLE_FILE_IMPORT_SUPPORTED_FORMATS = [
  'text/csv',
  'text/tab-separated-values',
  '.csv',
  '.tsv',
] as const
export const CANVAS_TABLE_IMPORT_MODEL = 'canvas-table-import'
export const CANVAS_TABLE_IMPORT_SUPPORTED_FORMATS = [
  'text/csv',
  'text/html',
  'text/markdown',
  'text/plain',
  'text/tab-separated-values',
] as const
export const CANVAS_TEXT_PASTE_SUPPORTED_FORMATS = [
  'text/html',
  'text/markdown',
  'text/plain',
] as const
export const CANVAS_TEXT_PASTE_IMPORT_MODEL = 'canvas-text-paste-import'
export const CANVAS_WHEEL_VIEWPORT_HORIZONTAL_PAN_MODIFIER = 'shift'
export const CANVAS_WHEEL_VIEWPORT_MODEL = 'canvas-wheel-viewport'
export const CANVAS_WHEEL_VIEWPORT_PAN_MODE = 'pan'
export const CANVAS_WHEEL_VIEWPORT_ZOOM_MODIFIER = 'meta'

export type CanvasClipboardCommand =
  | { kind: 'clone'; pasteIndex?: number }
  | { kind: 'copy'; pasteIndex?: number }
  | { kind: 'cut'; pasteIndex?: number }
  | { kind: 'duplicate'; pasteIndex?: number }
  | { kind: 'paste'; pasteIndex?: number }

export type CanvasClipboardCommandEffect<TItem extends { id: string }> =
  | { change: CanvasAppItemsChange<TItem>; type: 'items-change' }
  | { selection: readonly string[]; type: 'selection' }

export type CanvasClipboardCommandEffectContext<TItem extends { id: string }> =
  CanvasClipboardCommandExecutionContext<TItem>

export type CanvasClipboardCommandEffectPlanContext<
  TItem extends { id: string },
> = CanvasClipboardCommandExecutionContext<TItem>

export type CanvasClipboardCommandExecutionContext<
  TItem extends { id: string },
> = {
  commandAdapter?: unknown
  commitItemsChange: (
    change: CanvasAppItemsChange<TItem>,
    selection?: CanvasSelectionHistory,
  ) => boolean | undefined
  commitSelection: (action: CanvasSelectionAction) => boolean | undefined
  config?: unknown
  copyItemsToClipboard?: (selection: readonly string[]) => boolean
  createId: (prefix: string) => string
  getClipboardBounds?: (items: TItem[]) => Bounds | null
  getClipboardItems?: () => readonly TItem[]
  items: readonly TItem[]
  selection: readonly string[]
  setClipboardItems?: (items: readonly TItem[]) => boolean
  setEditing?: (editing: CanvasEditingAction) => void
  stageElement?: unknown
  viewport?: Viewport
}

export type CanvasClipboardCommandExecutionResult<TItem extends { id: string }> = {
  command: CanvasClipboardCommand
  effects: readonly CanvasClipboardCommandEffect<TItem>[]
  executed: boolean
}

export type CanvasAppItemsChange<TItem extends { id: string } = { id: string }> =
  | { items: TItem[]; type: 'add' }
  | { afterItems: TItem[]; beforeItems?: TItem[]; type: 'transform' }
  | { items: TItem[]; type: 'replace-changed' }
  | { groupId: string; selection: string[]; type: 'group-selection' }
  | { selection: string[]; type: 'remove-selection' }
  | { mode: CanvasCompatReorderMode; selection: string[]; type: 'reorder-selection' }
  | { selection: string[]; type: 'ungroup-selection' }
  | { type: 'resize-selection' }
  | { type: 'set-text' }

export type CanvasAppItemsChangeTransformer<TItem extends { id: string }> = {
  id: string
  transform: (input: {
    change: CanvasAppItemsChange<TItem>
    componentDefinitionRegistry?: unknown
    currentItems: readonly TItem[]
  }) => CanvasAppItemsChange<TItem>
}

export type CanvasAppHostItemsChangeCommitResult<
  TItem extends { id: string },
> = {
  committed: boolean
  transformedChange: CanvasAppItemsChange<TItem>
}

export type CommitCanvasAppHostItemsChangeArgs<TItem extends { id: string }> = {
  change: CanvasAppItemsChange<TItem>
  commitItemsChange: (change: CanvasAppItemsChange<TItem>) => boolean
  componentDefinitionRegistry?: unknown
  currentItems: readonly TItem[]
  itemsChangeTransformers?: readonly CanvasAppItemsChangeTransformer<TItem>[]
  selection?: CanvasSelectionHistory
}

export type CanvasStandardCommand =
  | { dx: number; dy: number; kind: 'nudge' }
  | { kind: 'align'; mode: string }
  | { kind: 'delete' }
  | { kind: 'distribute'; mode: string }
  | { kind: 'group' }
  | { kind: 'lock' }
  | { kind: 'redo' }
  | { kind: 'reorder'; mode: CanvasCompatReorderMode }
  | { kind: 'select-all' }
  | { kind: 'undo' }
  | { kind: 'ungroup' }
  | { kind: 'unlock-all' }

export type CanvasStandardCommandItemsChange<
  TItem extends { id: string },
> = CanvasAppItemsChange<TItem>

export type CanvasStandardCommandDocumentEffect<TItem extends { id: string }> =
  { change: CanvasStandardCommandItemsChange<TItem>; type: 'items-change' }

export type CanvasStandardCommandDocumentEffectContext<
  TItem extends { id: string },
> = CanvasStandardCommandExecutionContext<TItem>

export type CanvasStandardCommandEffectPlanContext<
  TItem extends { id: string },
> = CanvasStandardCommandExecutionContext<TItem>

export type CanvasStandardCommandExecutionContext<
  TItem extends { id: string },
> = {
  commandAdapter?: unknown
  commitItemsChange: (
    change: CanvasStandardCommandItemsChange<TItem>,
    selection?: CanvasSelectionHistory,
  ) => boolean | undefined
  commitSelection: (action: CanvasSelectionAction) => boolean
  config?: unknown
  createId: (prefix: string) => string
  items: readonly TItem[]
  redo?: () => void
  selection: readonly string[]
  setEditing?: (editing: CanvasEditingAction) => void
  setSelection?: (selection: string[]) => void
  undo?: () => void
}

type CanvasStandardCommandAdapterLike<TItem extends { id: string }> = {
  alignSelection?: (input: {
    items: TItem[]
    mode: string
    selection: string[]
  }) => TItem[]
  distributeSelection?: (input: {
    items: TItem[]
    mode: string
    selection: string[]
  }) => TItem[]
  lockSelection?: (input: {
    items: TItem[]
    selection: string[]
  }) => { items: TItem[]; selection: string[] }
  nudgeSelection?: (input: {
    dx: number
    dy: number
    items: TItem[]
    selection: string[]
  }) => TItem[]
  unlockAll?: (input: {
    items: TItem[]
    selection: string[]
  }) => { items: TItem[]; selection: string[] }
}

export type CanvasDataTransferImportRegistry<TAction, TScope extends string> = {
  resolvers: readonly CanvasDataTransferImportRegistryResolver<TAction, TScope>[]
}

export type CanvasDataTransferImportRegistryResolver<
  TAction,
  TScope extends string,
> = {
  id: string
  mode?: 'append' | 'exclusive'
  resolve: (input: { dataTransfer: DataTransfer | null }) =>
    | TAction
    | readonly TAction[]
    | null
    | undefined
  scope: TScope | readonly TScope[]
  supportedFormats?: readonly string[]
  title?: string
}

export type CanvasDataTransferImportRegistryResolverMetadata<
  TScope extends string,
> = {
  id: string
  scope: TScope | readonly TScope[]
  supportedFormats: readonly string[]
  title: string
}

export type CanvasKeyboardSelectionCycleDirection = 'next' | 'previous'

export type CanvasKeyboardSelectionCycleIntent<TId extends string = string> =
  | { kind: 'none' }
  | {
    direction: CanvasKeyboardSelectionCycleDirection
    fromId?: TId
    kind: 'cycle-selection'
    preventDefault?: boolean
    selectableIds: readonly TId[]
    targetId: TId
  }

export type CanvasKeyboardTextEditStartIntent<TId extends string = string> =
  | { kind: 'none' }
  | {
    initialText: string
    kind: 'start-text-edit'
    preventDefault?: boolean
    targetId: TId
  }

export type CanvasKeyboardTextFontSizeIntent = {
  delta: number
  kind: 'step-font-size'
  preventDefault?: boolean
}

export type CanvasKeyboardTextFontSizeKeyboardEvent = Pick<
  KeyboardEvent,
  'key' | 'metaKey' | 'ctrlKey' | 'shiftKey' | 'altKey'
>

export type CanvasPointerLaserInteraction = {
  kind: 'laser'
  laserTrail: CanvasLaserTrail
  pointerId?: number
  startScreen: Point
}

export type CanvasPointerPanInteraction = {
  kind: 'pan'
  pointerId?: number
  startScreen: Point
  viewport: Viewport
}

export type CanvasRichTextPasteSource =
  {
    format: 'text-html-rich' | 'text-markdown-rich'
    html?: string
    paragraphs: {
      align?: string
      bullet?: 'bullet' | 'numbered'
      headingLevel?: number
      lineHeight?: number
      runs: {
        bold?: boolean
        color?: string
        fontSize?: number
        italic?: boolean
        link?: string
        strikethrough?: boolean
        text: string
        underline?: boolean
      }[]
      spacingAfter?: number
      spacingBefore?: number
    }[]
    text: string
  }

export type CanvasTextPasteSource =
  | CanvasRichTextPasteSource
  | { format: 'text-plain'; text: string }

export type CanvasTableImportFormat =
  | 'text-csv'
  | 'text-delimited'
  | 'text-html'
  | 'text-markdown'
  | 'text-tsv'

export type CanvasTableImportSource = {
  format?: CanvasTableImportFormat
  name?: string
  rows: readonly (readonly string[])[]
}

export type CanvasTableImportTargetReplaceTarget = {
  id: string
  selection: readonly string[]
}

export type CanvasTableImportTargetReplaceRoute =
  | {
    intent: {
      rows: readonly (readonly string[])[]
      source: CanvasTableImportSource
      target: CanvasTableImportTargetReplaceTarget
    }
    kind: 'table-rows-replace'
  }
  | { kind: 'table-insert'; source: CanvasTableImportSource }

export type CanvasTextPasteReplaceTarget = {
  id: string
  selection: readonly string[]
}

export type CanvasTextPasteReplaceRoute =
  | {
    intent: {
      source: CanvasTextPasteSource
      target: CanvasTextPasteReplaceTarget
    }
    kind: 'text-replace'
  }
  | { kind: 'text-insert'; source: CanvasTextPasteSource }

export type RunCanvasClipboardCommand =
  (command: CanvasClipboardCommand) => readonly unknown[]

type CanvasSelectionHistory = {
  after?: readonly string[]
  before?: readonly string[]
}

type CanvasCompatReorderMode =
  | 'bringForward'
  | 'bringToFront'
  | 'sendBackward'
  | 'sendToBack'

type CanvasSelectionAction =
  | readonly string[]
  | ((selection: string[]) => string[])

type CanvasEditingAction =
  | { id: string; value: string }
  | null
  | ((editing: { id: string; value: string } | null) =>
      { id: string; value: string } | null)

type CanvasLaserTrail = {
  points: readonly Point[]
}

export type CanvasRichClipboardWriteInput =
  CanvasRichClipboardWriteInputBase & {
    extraItems?: Record<string, string>
  }

export function applyCanvasClipboardCommandEffect<TItem extends { id: string }>({
  context,
  effect,
}: {
  context: CanvasClipboardCommandExecutionContext<TItem>
  effect: CanvasClipboardCommandEffect<TItem>
}) {
  if (effect.type === 'items-change') {
    return context.commitItemsChange(effect.change)
  }

  return context.commitSelection([...effect.selection])
}

export function applyCanvasStandardDocumentEffect<TItem extends { id: string }>({
  context,
  effect,
}: {
  context: CanvasStandardCommandExecutionContext<TItem>
  effect: CanvasStandardCommandDocumentEffect<TItem>
}) {
  return context.commitItemsChange(effect.change)
}

export function commitCanvasAppHostItemsChange<TItem extends { id: string }>({
  change,
  commitItemsChange,
  componentDefinitionRegistry,
  currentItems,
  itemsChangeTransformers = [],
}: CommitCanvasAppHostItemsChangeArgs<TItem>):
  CanvasAppHostItemsChangeCommitResult<TItem> {
  const transformedChange = transformCanvasAppItemsChange({
    change,
    componentDefinitionRegistry,
    currentItems,
    transformers: itemsChangeTransformers,
  })

  return {
    committed: commitItemsChange(transformedChange),
    transformedChange,
  }
}

export function transformCanvasAppItemsChange<TItem extends { id: string }>({
  change,
  componentDefinitionRegistry,
  currentItems,
  transformers = [],
}: {
  change: CanvasAppItemsChange<TItem>
  componentDefinitionRegistry?: unknown
  currentItems: readonly TItem[]
  transformers?: readonly CanvasAppItemsChangeTransformer<TItem>[]
}): CanvasAppItemsChange<TItem> {
  return transformers.reduce(
    (currentChange, transformer) =>
      transformer.transform({
        change: currentChange,
        componentDefinitionRegistry,
        currentItems,
      }),
    change,
  )
}

export function createCanvasClipboardCommandEffectPlan() {
  return []
}

export function executeCanvasClipboardCommand<TItem extends { id: string }>({
  command,
  context,
}: {
  command: CanvasClipboardCommand
  context: CanvasClipboardCommandExecutionContext<TItem>
}): CanvasClipboardCommandExecutionResult<TItem> {
  const effects: CanvasClipboardCommandEffect<TItem>[] = []

  if (command.kind === 'copy') {
    context.copyItemsToClipboard?.(context.selection)
  }

  if (command.kind === 'cut') {
    context.copyItemsToClipboard?.(context.selection)
    effects.push({
      change: { selection: [...context.selection], type: 'remove-selection' },
      type: 'items-change',
    })
  }

  if (command.kind === 'duplicate') {
    const selected = context.items.filter((item) =>
      context.selection.includes(item.id))
    const duplicated = selected.map((item) => ({
      ...item,
      id: context.createId('item'),
      x: readCanvasNumberProperty(item, 'x') + 24,
      y: readCanvasNumberProperty(item, 'y') + 24,
    }))

    if (duplicated.length > 0) {
      effects.push({
        change: { items: duplicated as TItem[], type: 'add' },
        type: 'items-change',
      })
    }
  }

  const executed = effects.reduce(
    (committed, effect) =>
      applyCanvasClipboardCommandEffect({ context, effect }) === true ||
        committed,
    command.kind === 'copy',
  )

  return {
    command,
    effects,
    executed,
  }
}

export function executeCanvasStandardCommand<TItem extends { id: string }>({
  command,
  context,
}: {
  command: CanvasStandardCommand
  context: CanvasStandardCommandExecutionContext<TItem>
}) {
  const adapter = getCanvasStandardCommandAdapter<TItem>(context.commandAdapter)
  const items = [...context.items]
  const selection = [...context.selection]

  switch (command.kind) {
    case 'align': {
      const nextItems = adapter?.alignSelection?.({
        items,
        mode: command.mode,
        selection,
      })

      return nextItems
        ? context.commitItemsChange({ items: nextItems, type: 'replace-changed' })
        : false
    }
    case 'delete':
      return context.commitItemsChange(
        { selection, type: 'remove-selection' },
        { after: [], before: selection },
      )
    case 'distribute': {
      const nextItems = adapter?.distributeSelection?.({
        items,
        mode: command.mode,
        selection,
      })

      return nextItems
        ? context.commitItemsChange({ items: nextItems, type: 'replace-changed' })
        : false
    }
    case 'group':
      return context.commitItemsChange({
        groupId: context.createId('group'),
        selection,
        type: 'group-selection',
      })
    case 'lock': {
      const result = adapter?.lockSelection?.({
        items,
        selection,
      })

      return result
        ? context.commitItemsChange(
          { items: result.items, type: 'replace-changed' },
          { after: result.selection, before: selection },
        )
        : false
    }
    case 'nudge': {
      const nextItems = adapter?.nudgeSelection?.({
        dx: command.dx,
        dy: command.dy,
        items,
        selection,
      })

      return nextItems
        ? context.commitItemsChange({ items: nextItems, type: 'replace-changed' })
        : false
    }
    case 'reorder':
      return context.commitItemsChange({
        mode: command.mode,
        selection,
        type: 'reorder-selection',
      })
    case 'ungroup':
      return context.commitItemsChange({
        selection,
        type: 'ungroup-selection',
      })
    case 'redo':
      context.redo?.()
      return true
    case 'select-all':
      return context.commitSelection(context.items.map((item) => item.id))
    case 'undo':
      context.undo?.()
      return true
    case 'unlock-all': {
      const result = adapter?.unlockAll?.({
        items,
        selection,
      })

      return result
        ? context.commitItemsChange(
          { items: result.items, type: 'replace-changed' },
          { after: result.selection, before: selection },
        )
        : false
    }
  }
}

function getCanvasStandardCommandAdapter<TItem extends { id: string }>(
  adapter: unknown,
): CanvasStandardCommandAdapterLike<TItem> | null {
  return adapter && typeof adapter === 'object'
    ? adapter as CanvasStandardCommandAdapterLike<TItem>
    : null
}

export function copyCanvasClipboardSelection({
  pasteIndex = 0,
  runClipboardCommand,
}: {
  pasteIndex?: number
  runClipboardCommand: RunCanvasClipboardCommand
}) {
  return runClipboardCommand({ kind: 'copy', pasteIndex })
}

export function cutCanvasClipboardSelection({
  pasteIndex = 0,
  runClipboardCommand,
}: {
  pasteIndex?: number
  runClipboardCommand: RunCanvasClipboardCommand
}) {
  return runClipboardCommand({ kind: 'cut', pasteIndex })
}

export function duplicateCanvasClipboardSelection({
  pasteIndex = 0,
  runClipboardCommand,
}: {
  pasteIndex?: number
  runClipboardCommand: RunCanvasClipboardCommand
  selection?: readonly string[]
}) {
  return runClipboardCommand({ kind: 'duplicate', pasteIndex })
}

export function pasteCanvasClipboardSelection({
  pasteIndex = 0,
  runClipboardCommand,
}: {
  pasteIndex?: number
  runClipboardCommand: RunCanvasClipboardCommand
}) {
  return runClipboardCommand({ kind: 'paste', pasteIndex })
}

export function getCanvasExternalClipboardPasteCommandRoute({
  hasInternalClipboard,
}: {
  hasInternalClipboard: boolean
}) {
  return hasInternalClipboard ? 'internal-clipboard' : 'external-clipboard'
}

export function createCanvasDataTransferImportRegistry<
  TAction,
  TScope extends string,
>({
  resolvers,
}: {
  resolvers: readonly CanvasDataTransferImportRegistryResolver<TAction, TScope>[]
}): CanvasDataTransferImportRegistry<TAction, TScope> {
  return { resolvers }
}

export function createCanvasDataTransferImportActionPlan<TAction>({
  resolvers,
}: {
  resolvers: readonly {
    resolve: () => TAction | readonly TAction[] | null | undefined
  }[]
}): TAction[] {
  return resolvers.flatMap((resolver) => normalizeCanvasActionResult(
    resolver.resolve(),
  ))
}

export function createCanvasDataTransferImportActionPlanFromRegistry<
  TAction,
  TScope extends string,
>({
  dataTransfer,
  registry,
  scope,
}: {
  dataTransfer: DataTransfer | null
  registry: CanvasDataTransferImportRegistry<TAction, TScope>
  scope: TScope
}): TAction[] {
  const actions: TAction[] = []

  for (const resolver of registry.resolvers) {
    if (!isCanvasDataTransferImportScopeMatch(resolver.scope, scope)) {
      continue
    }

    const next = normalizeCanvasActionResult(resolver.resolve({ dataTransfer }))

    actions.push(...next)

    if (resolver.mode === 'exclusive' && next.length > 0) {
      break
    }
  }

  return actions
}

export function getCanvasDataTransferImportRegistryMetadata<
  TAction,
  TScope extends string,
>({
  registry,
  scope,
}: {
  registry: CanvasDataTransferImportRegistry<TAction, TScope>
  scope: TScope
}): CanvasDataTransferImportRegistryResolverMetadata<TScope>[] {
  return registry.resolvers
    .filter((resolver) =>
      isCanvasDataTransferImportScopeMatch(resolver.scope, scope))
    .map((resolver) => ({
      id: resolver.id,
      scope: resolver.scope,
      supportedFormats: resolver.supportedFormats ?? [],
      title: resolver.title ?? resolver.id,
    }))
}

export function runCanvasDataTransferImportActionPlan<TAction>({
  actions,
  onConsumed,
  runAction,
}: {
  actions: readonly TAction[]
  onConsumed?: () => void
  runAction: (action: TAction) => boolean | void
}) {
  for (const action of actions) {
    if (runAction(action)) {
      onConsumed?.()

      return { action, consumed: true }
    }
  }

  return { action: null, consumed: false }
}

export async function createCanvasExternalClipboardPasteActionPlan<TAction>({
  resolvers,
}: {
  resolvers: readonly {
    resolve: () =>
      | TAction
      | readonly TAction[]
      | null
      | undefined
      | Promise<TAction | readonly TAction[] | null | undefined>
  }[]
}) {
  const results = await Promise.all(resolvers.map((resolver) =>
    resolver.resolve()))

  return results.flatMap((result) => normalizeCanvasActionResult(result))
}

export function createCanvasExternalClipboardImagePasteActionResolver<TSource>({
  createAction,
  getSource,
  readImageSource,
}: {
  createAction?: (source: TSource) => TSource
  getSource?: () => TSource | null | undefined | Promise<TSource | null | undefined>
  pasteSource?: (source: TSource) => boolean | Promise<boolean>
  readImageSource?: () => TSource | null | undefined | Promise<TSource | null | undefined>
}) {
  return {
    resolve: async () => {
      const source = await (getSource ?? readImageSource)?.()

      return source ? createAction?.(source) ?? source : null
    },
  }
}

export function readCanvasDataTransferTextCandidate<
  TCandidate extends string,
>({
  candidates,
  dataTransfer,
}: {
  candidates: readonly TCandidate[]
  dataTransfer: { getData?: (format: string) => string } | null
}) {
  if (!dataTransfer) {
    return null
  }

  for (const candidate of candidates) {
    const text = dataTransfer.getData?.(candidate) ?? ''

    if (text) {
      return {
        candidate,
        rawText: text,
        text,
      }
    }
  }

  return null
}

export function readCanvasDataTransferJSONCandidate<
  TSource,
  TCandidate extends string | { mimeType: string },
>({
  candidates,
  dataTransfer,
  extractTextJSON = false,
  parseValue,
}: {
  candidates: readonly TCandidate[]
  dataTransfer: { getData?: (format: string) => string } | null
  extractTextJSON?: boolean
  parseValue: (input: {
    candidate: TCandidate
    json: unknown
    jsonText?: string
    rawText: string
  }) => TSource | null
}): { value: TSource } | null {
  if (!dataTransfer) {
    return null
  }

  for (const candidate of candidates) {
    const mimeType = typeof candidate === 'string'
      ? candidate
      : candidate.mimeType
    const rawText = dataTransfer.getData?.(mimeType) ?? ''
    const jsonTexts = [
      rawText,
      ...(extractTextJSON ? extractCanvasJSONSnippets(rawText) : []),
    ].filter(Boolean)

    for (const jsonText of jsonTexts) {
      try {
        const json = JSON.parse(jsonText)
        const parsed = parseValue({
          candidate,
          json,
          jsonText,
          rawText,
        })

        if (parsed) {
          return { value: parsed }
        }
      } catch {
        continue
      }
    }
  }

  return null
}

export function getCanvasKeyboardSelectionCycleIntent<TId extends string>({
  event,
  selectableIds,
  selection,
}: {
  event: Partial<Pick<
    KeyboardEvent,
    'altKey' | 'ctrlKey' | 'key' | 'metaKey' | 'shiftKey'
  >> & { target?: EventTarget | null }
  selectableIds: readonly TId[]
  selection: readonly TId[]
  targetSelectors?: string
}): CanvasKeyboardSelectionCycleIntent<TId> {
  if (event.key !== 'Tab' || selectableIds.length === 0) {
    return { kind: 'none' }
  }

  const direction = event.shiftKey ? 'previous' : 'next'
  const selectedIndex = selection[0]
    ? selectableIds.indexOf(selection[0])
    : -1
  const offset = direction === 'next' ? 1 : -1
  const targetIndex =
    (selectedIndex + offset + selectableIds.length) % selectableIds.length
  const targetId = selectableIds[targetIndex]

  return {
    direction,
    fromId: selection[0],
    kind: 'cycle-selection',
    preventDefault: true,
    selectableIds,
    targetId,
  }
}

export function getCanvasKeyboardTextEditStartIntent<TId extends string>({
  event,
  selection,
}: {
  blockedTargetSelectors?: string
  event: Partial<Pick<
    KeyboardEvent,
    'altKey' | 'ctrlKey' | 'key' | 'metaKey'
  >> & { target?: EventTarget | null }
  isEditableTextSelection?: boolean
  isReservedShortcut?: boolean | (() => boolean)
  selection: readonly TId[]
}): CanvasKeyboardTextEditStartIntent<TId> {
  const targetId = selection[0]

  if (!targetId || selection.length !== 1) {
    return { kind: 'none' }
  }

  if (
    event.key &&
    event.key.length === 1 &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.altKey
  ) {
    return {
      initialText: event.key === ' ' ? '' : event.key,
      kind: 'start-text-edit',
      preventDefault: true,
      targetId,
    }
  }

  if (event.key === 'Enter') {
    return {
      initialText: '',
      kind: 'start-text-edit',
      preventDefault: true,
      targetId,
    }
  }

  return { kind: 'none' }
}

export function getCanvasKeyboardTextFontSizeShortcutIntent({
  event,
}: {
  event: CanvasKeyboardTextFontSizeKeyboardEvent
}): CanvasKeyboardTextFontSizeIntent | { kind: 'none' } {
  const mod = event.metaKey || event.ctrlKey

  if (!mod || event.altKey) {
    return { kind: 'none' }
  }

  if (event.shiftKey && (event.key === '>' || event.key === '.')) {
    return {
      delta: CANVAS_KEYBOARD_TEXT_FONT_SIZE_STEP,
      kind: 'step-font-size',
      preventDefault: true,
    }
  }

  if (event.shiftKey && (event.key === '<' || event.key === ',')) {
    return {
      delta: -CANVAS_KEYBOARD_TEXT_FONT_SIZE_STEP,
      kind: 'step-font-size',
      preventDefault: true,
    }
  }

  return { kind: 'none' }
}

export function getCanvasImageInsertCenter({
  event,
  stageElement,
  viewport,
}: CanvasInsertPositionInput): Point {
  return getCanvasInsertPosition({ event, stageElement, viewport })
}

export function getCanvasMediaInsertPosition({
  event,
  stageElement,
  viewport,
}: CanvasInsertPositionInput): Point {
  return getCanvasInsertPosition({ event, stageElement, viewport })
}

export function getCanvasTableInsertCenter({
  event,
  stageElement,
  viewport,
}: CanvasInsertPositionInput): Point {
  return getCanvasInsertPosition({ event, stageElement, viewport })
}

export function getCanvasTextPasteInsertPosition({
  event,
  stageElement,
  viewport,
}: CanvasInsertPositionInput): Point {
  return getCanvasInsertPosition({ event, stageElement, viewport })
}

type CanvasInsertPositionInput = {
  event?: { clientX: number; clientY: number }
  stageElement: CanvasAppStageElement
  viewport: Viewport
}

export function getNextCanvasDrawingPoints(input: {
  currentWorld: Point
  points: readonly Point[]
  shiftKey?: boolean
  startWorld?: Point
}): Point[]
export function getNextCanvasDrawingPoints(
  points: readonly Point[],
  point: Point,
): Point[]
export function getNextCanvasDrawingPoints(
  input: readonly Point[] | {
    currentWorld: Point
    points: readonly Point[]
    shiftKey?: boolean
    startWorld?: Point
  },
  point?: Point,
) {
  if (!('points' in input)) {
    if (!point) {
      return [...input]
    }

    const last = input.at(-1)

    return !last || getCanvasPointDistance(last, point) >= 2
      ? [...input, point]
      : [...input]
  }

  const points = input.points
  const last = points.at(-1)

  if (!last || getCanvasPointDistance(last, input.currentWorld) >= 2) {
    return [...points, input.currentWorld]
  }

  return [...points]
}

export function getCanvasEraserHitItemIds({
  itemReadModel,
  points,
  radius = 8,
}: {
  itemReadModel: {
    getAllItems: () => readonly CanvasEraserHitItem[]
    [key: string]: unknown
  }
  points: readonly Point[]
  radius?: number
  scene?: unknown
}) {
  const hitIds = new Set<string>()

  for (const item of itemReadModel.getAllItems()) {
    const itemPoints = item.points ?? []

    if (itemPoints.some((itemPoint) =>
      points.some((point) => getCanvasPointDistance(itemPoint, point) <= radius)
    )) {
      hitIds.add(item.id)
    }
  }

  return [...hitIds]
}

export function getCanvasMergedEraserHitIds(
  current: Iterable<string>,
  next: Iterable<string>,
) {
  return [...new Set([...current, ...next])]
}

export function getCanvasRichTextPasteSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
): CanvasRichTextPasteSource | null {
  const html = dataTransfer?.getData('text/html') ?? ''

  if (html.trim()) {
    const text = stripCanvasHTMLText(html)

    return {
      format: 'text-html-rich',
      html,
      paragraphs: [{
        runs: [{ text }],
      }],
      text,
    }
  }

  return null
}

export function getCanvasTextPasteSourcesFromDataTransfer(
  dataTransfer: DataTransfer | null,
): CanvasTextPasteSource[] {
  if (!dataTransfer) {
    return []
  }

  const sources: CanvasTextPasteSource[] = []
  const rich = getCanvasRichTextPasteSourceFromDataTransfer(dataTransfer)
  const text = dataTransfer.getData('text/plain')

  if (rich) {
    sources.push(rich)
  }

  if (text.trim()) {
    sources.push({ format: 'text-plain', text })
  }

  return sources
}

export const getCanvasTextPasteSourceCandidatesFromDataTransfer =
  getCanvasTextPasteSourcesFromDataTransfer

export function getCanvasTextPasteSourceText(source: CanvasTextPasteSource) {
  return source.text
}

export function createCanvasTextPasteItems<TItem>({
  createId,
  importers,
  position,
  text,
  viewport,
}: {
  createId: (prefix: string) => string
  importers: readonly {
    createItems: (input: {
      createId: (prefix: string) => string
      position: Point
      text: string
      viewport: Viewport
    }) => TItem[] | null
    id: string
  }[]
  position: Point
  text: string
  viewport: Viewport
}) {
  for (const importer of importers) {
    const items = importer.createItems({
      createId,
      position,
      text,
      viewport,
    })

    if (items && items.length > 0) {
      return {
        importerId: importer.id,
        items,
      }
    }
  }

  return null
}

export function routeCanvasTextPasteReplace({
  getTarget,
  selection,
  source,
}: {
  getTarget: (input: { selection: readonly string[] }) =>
    CanvasTextPasteReplaceTarget | null
  selection: readonly string[]
  source: CanvasTextPasteSource
}): CanvasTextPasteReplaceRoute {
  const target = getTarget({ selection })

  return target
    ? {
        intent: {
          source,
          target,
        },
        kind: 'text-replace',
      }
    : {
        kind: 'text-insert',
        source,
      }
}

export function getCanvasTableColumnCount(
  rows: readonly (readonly string[])[],
) {
  return rows.reduce((count, row) => Math.max(count, row.length), 0)
}

export function getCanvasTableComponentSize({
  columnCount,
  rowCount,
}: {
  columnCount: number
  rowCount: number
}, options: {
  cellSize?: Partial<{ h: number; w: number }>
  maxSize?: Partial<{ h: number; w: number }>
  minSize?: Partial<{ h: number; w: number }>
} = {}) {
  const width = columnCount * (options.cellSize?.w ?? 140)
  const height = rowCount * (options.cellSize?.h ?? 44)

  return {
    h: clampCanvasNumber(
      height,
      options.minSize?.h ?? 80,
      options.maxSize?.h ?? 520,
    ),
    w: clampCanvasNumber(
      width,
      options.minSize?.w ?? 240,
      options.maxSize?.w ?? 920,
    ),
  }
}

export function getCanvasTableFileFromList(files: FileList | null) {
  return getCanvasTableFilesFromList(files)[0] ?? null
}

export function getCanvasTableFilesFromList(files: FileList | null) {
  return Array.from(files ?? []).filter(isCanvasTableFile)
}

export function getCanvasTableFileFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getCanvasTableFileFromList(dataTransfer?.files ?? null)
}

export function getCanvasTableFilesFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  return getCanvasTableFilesFromList(dataTransfer?.files ?? null)
}

export function getCanvasTableSourceFromDataTransfer(
  dataTransfer: DataTransfer | null,
) {
  if (!dataTransfer) {
    return null
  }

  return getCanvasTableSourceFromHTML(dataTransfer.getData('text/html')) ??
    getCanvasTableSourceFromText(dataTransfer.getData('text/csv'), {
      format: 'text-csv',
    }) ??
    getCanvasTableSourceFromText(
      dataTransfer.getData('text/tab-separated-values'),
      { format: 'text-tsv' },
    ) ??
    getCanvasTableSourceFromText(dataTransfer.getData('text/plain'))
}

export function getCanvasTableSourceFromHTML(
  value: string,
): CanvasTableImportSource | null {
  if (!value || typeof DOMParser === 'undefined') {
    return null
  }

  const doc = new DOMParser().parseFromString(value, 'text/html')
  const table = doc.querySelector('table')

  if (!table) {
    return null
  }

  const rows = Array.from(table.querySelectorAll('tr')).map((row) =>
    Array.from(row.querySelectorAll('th,td')).map((cell) =>
      normalizeCanvasCellText(cell.textContent ?? '')))

  return rows.length > 0
    ? {
        format: 'text-html',
        rows: normalizeCanvasTableRows(rows),
      }
    : null
}

export function getCanvasTableSourceFromText(
  text: string,
  options: { format?: CanvasTableImportFormat; name?: string } = {},
): CanvasTableImportSource | null {
  const trimmed = text.trim()

  if (!trimmed) {
    return null
  }

  const format = options.format ?? inferCanvasTableTextFormat(trimmed)
  const rows = parseCanvasTableText(trimmed, format)

  return getCanvasTableColumnCount(rows) >= 2 && rows.length >= 2
    ? {
        format,
        name: options.name,
        rows: normalizeCanvasTableRows(rows),
      }
    : null
}

export async function readCanvasTableFileSource(
  file: Blob & { name?: string },
): Promise<CanvasTableImportSource | null> {
  const text = await file.text()
  const format = file.name?.toLowerCase().endsWith('.csv')
    ? 'text-csv'
    : file.name?.toLowerCase().endsWith('.tsv')
      ? 'text-tsv'
      : undefined

  return getCanvasTableSourceFromText(text, {
    format,
    name: file.name,
  })
}

export async function readCanvasTableFileSources(
  files: readonly (Blob & { name?: string })[],
) {
  const sources = await Promise.all(files.map(readCanvasTableFileSource))

  return sources.filter((source): source is CanvasTableImportSource =>
    source !== null)
}

export function normalizeCanvasTableRows(
  rows: readonly (readonly string[])[],
  options: {
    fallbackRows?: readonly (readonly string[])[]
    maxCellLength?: number
    maxColumns?: number
    maxRows?: number
  } = {},
) {
  const maxRows = options.maxRows ?? 100
  const maxColumns = options.maxColumns ?? 20
  const maxCellLength = options.maxCellLength ?? 300
  const normalized = rows
    .slice(0, maxRows)
    .map((row) =>
      row
        .slice(0, maxColumns)
        .map((cell) => normalizeCanvasCellText(cell).slice(0, maxCellLength)))
    .filter((row) => row.some(Boolean))

  return normalized.length > 0
    ? normalized
    : (options.fallbackRows ?? []).map((row) => [...row])
}

export function routeCanvasTableImportTargetReplace({
  getTarget,
  normalizeRows,
  selection,
  source,
}: {
  getTarget: (input: { selection: readonly string[] }) =>
    CanvasTableImportTargetReplaceTarget | null
  normalizeRows: (input: { source: CanvasTableImportSource }) =>
    readonly (readonly string[])[]
  selection: readonly string[]
  source: CanvasTableImportSource
}): CanvasTableImportTargetReplaceRoute {
  const target = getTarget({ selection })

  return target
    ? {
        intent: {
          rows: normalizeRows({ source }),
          source,
          target,
        },
        kind: 'table-rows-replace',
      }
    : {
        kind: 'table-insert',
        source,
      }
}

export function runCanvasWheelViewport({
  config,
  event,
  rect,
  setViewport,
}: {
  config: Parameters<typeof getCanvasWheelViewport>[0]['config']
  event: WheelEvent
  rect: DOMRect | { left: number; top: number }
  setViewport: (action: (viewport: Viewport) => Viewport) => void
}) {
  setViewport((viewport) =>
    getCanvasWheelViewport({
      config,
      input: event,
      point: {
        x: event.clientX - rect.left,
        y: event.clientY - rect.top,
      },
      viewport,
    }) ?? viewport)
  event.preventDefault()
}

export function startCanvasPointerPanInteraction({
  input,
  startScreen,
  viewport,
}: {
  input?: { pointerId?: number }
  startScreen: Point
  viewport: Viewport
}) {
  return {
    interaction: {
      kind: 'pan' as const,
      pointerId: input?.pointerId,
      startScreen,
      viewport,
    },
    kind: 'interaction' as const,
  }
}

export function previewCanvasPointerPanInteraction({
  currentScreen,
  interaction,
}: {
  config?: unknown
  currentScreen: Point
  interaction: CanvasPointerPanInteraction
}) {
  return {
    kind: 'preview' as const,
    viewport: {
      ...interaction.viewport,
      x: interaction.viewport.x + currentScreen.x - interaction.startScreen.x,
      y: interaction.viewport.y + currentScreen.y - interaction.startScreen.y,
    },
  }
}

export function startCanvasPointerLaserInteraction({
  input,
  startScreen,
  startWorld,
}: {
  config?: unknown
  input?: { pointerId?: number }
  pointerGesture?: string
  startScreen: Point
  startWorld: Point
}) {
  const laserTrail = { points: [startWorld] }

  return {
    interaction: {
      kind: 'laser' as const,
      laserTrail,
      pointerId: input?.pointerId,
      startScreen,
    },
    kind: 'interaction' as const,
    laserTrail,
  }
}

export function previewCanvasPointerLaserInteraction({
  currentWorld,
  interaction,
}: {
  config?: unknown
  currentScreen: Point
  currentWorld: Point
  interaction: CanvasPointerLaserInteraction
}) {
  const points = [...interaction.laserTrail.points, currentWorld].slice(-64)
  const laserTrail = { points }

  return {
    interaction: {
      ...interaction,
      laserTrail,
    },
    kind: 'preview' as const,
    laserTrail,
  }
}

export function useCanvasAppStageElement(): CanvasAppStageElement {
  const elementRef = useRef<HTMLElement | null>(null)

  return useMemo(() => ({
    addWheelListener(listener: (event: WheelEvent, rect: CanvasAppStageRect) => void) {
      const element = elementRef.current

      if (!element) {
        return () => {}
      }

      const onWheel = (event: WheelEvent) => {
        const rect = element.getBoundingClientRect()

        listener(event, {
          height: rect.height,
          left: rect.left,
          top: rect.top,
          width: rect.width,
        })
      }

      element.addEventListener('wheel', onWheel, { passive: false })

      return () => element.removeEventListener('wheel', onWheel)
    },
    capturePointer(pointerId: number) {
      elementRef.current?.setPointerCapture?.(pointerId)
    },
    getElement: () => elementRef.current,
    getRect: () => elementRef.current?.getBoundingClientRect() ?? null,
    getScreenPoint: (event: { clientX: number; clientY: number }) => {
      const rect = elementRef.current?.getBoundingClientRect()

      return rect
        ? {
            x: event.clientX - rect.left,
            y: event.clientY - rect.top,
          }
        : { x: event.clientX, y: event.clientY }
    },
    getViewportCenter: (viewport: Viewport) => {
      const rect = elementRef.current?.getBoundingClientRect()

      return rect
        ? {
            x: (rect.width / 2 - viewport.x) / viewport.scale,
            y: (rect.height / 2 - viewport.y) / viewport.scale,
          }
        : null
    },
    mount: {
      ref: (element: HTMLElement | null) => {
        elementRef.current = element
      },
    },
    releasePointer(pointerId: number) {
      elementRef.current?.releasePointerCapture?.(pointerId)
    },
  }), [])
}

export type CanvasAppStageElement = {
  addWheelListener: (
    listener: (event: WheelEvent, rect: CanvasAppStageRect) => void,
  ) => () => void
  capturePointer: (pointerId: number) => void
  getElement: () => HTMLElement | null
  getRect: () => DOMRect | null
  getScreenPoint: (event: { clientX: number; clientY: number }) => Point
  getViewportCenter: (viewport: Viewport) => Point | null
  mount: {
    ref: (element: HTMLElement | null) => void
  }
  releasePointer: (pointerId: number) => void
}

export type CanvasAppStageRect = {
  height: number
  left: number
  top: number
  width: number
}

type CanvasEraserHitItem = {
  h: number
  id: string
  opacity?: number
  points?: readonly Point[]
  stroke?: string
  strokeWidth?: number
  type: 'highlight' | 'marker'
  w: number
  x: number
  y: number
}

export async function writeCanvasRichClipboardPayload({
  extraItems,
  ...input
}: CanvasRichClipboardWriteInput): Promise<CanvasRichClipboardWriteMode> {
  void extraItems

  return writeCanvasRichClipboardPayloadBase(input)
}

function getCanvasInsertPosition({
  event,
  stageElement,
  viewport,
}: CanvasInsertPositionInput): Point {
  if (event) {
    const point = stageElement.getScreenPoint(event)

    return {
      x: (point.x - viewport.x) / viewport.scale,
      y: (point.y - viewport.y) / viewport.scale,
    }
  }

  return stageElement.getViewportCenter(viewport) ?? { x: 0, y: 0 }
}

function normalizeCanvasActionResult<TAction>(
  value: TAction | readonly TAction[] | null | undefined,
): TAction[] {
  if (!value) {
    return []
  }

  if (Array.isArray(value)) {
    return [...value] as TAction[]
  }

  return [value as TAction]
}

function isCanvasDataTransferImportScopeMatch<TScope extends string>(
  scope: TScope | readonly TScope[],
  expected: TScope,
) {
  return Array.isArray(scope) ? scope.includes(expected) : scope === expected
}

function extractCanvasJSONSnippets(value: string) {
  return Array.from(value.matchAll(/<script[^>]*type=["']application\/json["'][^>]*>([\s\S]*?)<\/script>/gi))
    .map((match) => match[1] ?? '')
}

function readCanvasNumberProperty(value: unknown, key: string) {
  return isCanvasRecord(value) && typeof value[key] === 'number'
    ? value[key]
    : 0
}

function isCanvasRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function getCanvasPointDistance(a: Point, b: Point) {
  return Math.hypot(a.x - b.x, a.y - b.y)
}

function stripCanvasHTMLText(value: string) {
  if (typeof DOMParser === 'undefined') {
    return value.replace(/<[^>]+>/g, ' ')
  }

  return new DOMParser().parseFromString(value, 'text/html').body.textContent ??
    ''
}

function isCanvasTableFile(file: File) {
  const name = file.name.toLowerCase()

  return file.type === 'text/csv' ||
    file.type === 'text/tab-separated-values' ||
    name.endsWith('.csv') ||
    name.endsWith('.tsv')
}

function inferCanvasTableTextFormat(text: string): CanvasTableImportFormat {
  if (text.includes('\t')) {
    return 'text-tsv'
  }

  if (text.includes(',')) {
    return 'text-csv'
  }

  return 'text-delimited'
}

function parseCanvasTableText(
  text: string,
  format: CanvasTableImportFormat,
) {
  if (format === 'text-csv') {
    return text.split(/\r?\n/).map(parseCanvasCSVLine)
  }

  if (format === 'text-tsv') {
    return text.split(/\r?\n/).map((line) => line.split('\t'))
  }

  if (format === 'text-markdown') {
    return text
      .split(/\r?\n/)
      .filter((line) => line.includes('|') && !/^\s*\|?\s*:?-{3,}/.test(line))
      .map((line) => line.split('|').map((cell) => cell.trim()).filter(Boolean))
  }

  return text.split(/\r?\n/).map((line) => line.split(/\s{2,}|\t|,/))
}

function parseCanvasCSVLine(line: string) {
  const cells: string[] = []
  let cell = ''
  let quoted = false

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]

    if (char === '"' && line[index + 1] === '"') {
      cell += '"'
      index += 1
    } else if (char === '"') {
      quoted = !quoted
    } else if (char === ',' && !quoted) {
      cells.push(cell)
      cell = ''
    } else {
      cell += char
    }
  }

  cells.push(cell)

  return cells
}

function normalizeCanvasCellText(value: string) {
  return value.replace(/\s+/g, ' ').trim()
}

function clampCanvasNumber(value: number, min: number, max: number) {
  return Math.max(min, Math.min(max, value))
}
