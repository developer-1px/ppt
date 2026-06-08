import { useMemo, useState } from 'react'
import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import {
  reducePatternData,
  toolbarDefinition,
  useToolbarPattern,
} from '@interactive-os/aria/react'
import type {
  PatternData,
  PatternEvent,
  PatternOptions,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'

type PatternButtonProps = ButtonHTMLAttributes<HTMLButtonElement>
type PatternDivProps = HTMLAttributes<HTMLDivElement>
type PatternElementProps = HTMLAttributes<HTMLElement>
type ToolbarOrientation = NonNullable<PatternOptions['orientation']>

const EMPTY_TOOLBAR_KEYS: readonly never[] = []

export type ActionToolbarItem<TKey extends string> = {
  action: TKey
  label: string
}

export function useManagedToolbarPattern<TKey extends string>({
  activeKey,
  disabledKeys = EMPTY_TOOLBAR_KEYS,
  elementIdPrefix,
  handlers,
  items,
  label,
  omitPressed = false,
  orientation = 'horizontal',
  rootKeys,
  selectedKeys,
}: {
  activeKey?: TKey | null
  disabledKeys?: readonly TKey[]
  elementIdPrefix: string
  handlers: Partial<Record<TKey, () => void>>
  items: Record<TKey, { label: string }>
  label: string
  omitPressed?: boolean
  orientation?: ToolbarOrientation
  rootKeys: readonly TKey[]
  selectedKeys?: readonly TKey[]
}) {
  const [activeToolbarKey, setActiveToolbarKey] = useState<TKey | null>(
    activeKey ?? null,
  )

  const data = useMemo(() =>
    toolbarPatternData<TKey>({
      activeKey: activeToolbarKey,
      disabledKeys,
      items,
      label,
      rootKeys,
      selectedKeys,
    }), [activeToolbarKey, disabledKeys, items, label, rootKeys, selectedKeys])
  const toolbar = useToolbarPattern(
    data,
    (event) => {
      const nextActiveKey = nextToolbarActiveKey<TKey>(
        data,
        event,
        disabledKeys,
      )

      if (nextActiveKey) {
        setActiveToolbarKey(nextActiveKey)
      }

      handleToolbarSelection<TKey>(event, handlers)
    },
    {
      elementIdPrefix,
      orientation,
    },
  )

  return {
    itemProps: toolbarItemPropsByKey<TKey>(toolbar.renderItems, {
      omitPressed,
    }),
    rootProps: toolbar.rootProps,
  }
}

export function useActionToolbarPattern<TKey extends string>({
  actions,
  activeKey,
  disabledKeys,
  elementIdPrefix,
  label,
  omitPressed = true,
  onSelect,
  orientation,
  selectedKeys,
}: {
  actions: readonly ActionToolbarItem<TKey>[]
  activeKey?: TKey | null
  disabledKeys?: readonly TKey[]
  elementIdPrefix: string
  label: string
  omitPressed?: boolean
  onSelect: (action: TKey) => void
  orientation?: ToolbarOrientation
  selectedKeys?: readonly TKey[]
}) {
  const rootKeys = useMemo(() => actionToolbarKeys(actions), [actions])
  const items = useMemo(() => actionToolbarItems(actions), [actions])
  const handlers = useMemo(() =>
    actionToolbarHandlers(actions, onSelect), [actions, onSelect])

  return useManagedToolbarPattern<TKey>({
    activeKey,
    disabledKeys,
    elementIdPrefix,
    handlers,
    items,
    label,
    omitPressed,
    orientation,
    rootKeys,
    selectedKeys,
  })
}

export function patternButtonProps(
  props: PatternElementProps,
): PatternButtonProps {
  return props as PatternButtonProps
}

export function patternDivProps(props: PatternElementProps): PatternDivProps {
  return props as PatternDivProps
}

export function selectedToolbarKeys(event: PatternEvent) {
  return event.type === 'select' ? event.keys : []
}

export function nextToolbarActiveKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
  disabledKeys: readonly TKey[] = [],
) {
  const nextKey = nextToolbarEventKey(data, event)

  if (!nextKey || disabledKeys.includes(nextKey as TKey)) {
    return null
  }

  return nextKey as TKey
}

export function disabledToolbarKeys<TKey extends string>(
  candidates: readonly (readonly [key: TKey, disabled: boolean])[],
) {
  return candidates.flatMap(([key, disabled]) => (disabled ? [key] : []))
}

export function handleToolbarSelection<TKey extends string>(
  event: PatternEvent,
  handlers: Partial<Record<TKey, () => void>>,
) {
  for (const key of selectedToolbarKeys(event)) {
    handlers[key as TKey]?.()
  }
}

export function toolbarPatternData<TKey extends string>({
  activeKey,
  disabledKeys = [],
  items,
  label,
  rootKeys,
  selectedKeys,
}: {
  activeKey?: TKey | null
  disabledKeys?: readonly TKey[]
  items: Record<TKey, { label: string }>
  label: string
  rootKeys: readonly TKey[]
  selectedKeys?: readonly TKey[]
}): PatternData {
  return {
    items,
    relations: { rootKeys },
    refs: { label },
    state: {
      activeKey: enabledToolbarActiveKey(
        activeKey,
        rootKeys,
        disabledKeys,
      ),
      ...(disabledKeys.length > 0 ? { disabledKeys } : {}),
      ...(selectedKeys ? { selectedKeys } : {}),
    },
  }
}

export function toolbarItemPropsByKey<TKey extends string>(
  items: readonly ReactToolbarRenderItem[],
  options: { omitPressed?: boolean } = {},
) {
  return Object.fromEntries(
    items.map((item) => {
      const props = patternButtonProps(item.itemProps)

      return [
        item.key,
        options.omitPressed ? omitAriaPressed(props) : props,
      ]
    }),
  ) as Record<TKey, PatternButtonProps>
}

export function firstEnabledToolbarKey<TKey extends string>(
  keys: readonly TKey[],
  disabledKeys: readonly TKey[],
) {
  return keys.find((key) => !disabledKeys.includes(key)) ?? keys[0] ?? null
}

function enabledToolbarActiveKey<TKey extends string>(
  activeKey: TKey | null | undefined,
  rootKeys: readonly TKey[],
  disabledKeys: readonly TKey[],
) {
  if (activeKey && !disabledKeys.includes(activeKey)) {
    return activeKey
  }

  return firstEnabledToolbarKey(rootKeys, disabledKeys)
}

function nextToolbarEventKey(data: PatternData, event: PatternEvent) {
  if (event.type === 'focus') {
    return event.key
  }

  if (event.type === 'navigate') {
    return reducePatternData(toolbarDefinition, data, event).state?.activeKey ?? null
  }

  return selectedToolbarKeys(event)[0] ?? null
}

function actionToolbarKeys<TKey extends string>(
  actions: readonly ActionToolbarItem<TKey>[],
): readonly TKey[] {
  return actions.map(({ action }) => action)
}

function actionToolbarItems<TKey extends string>(
  actions: readonly ActionToolbarItem<TKey>[],
): Record<TKey, { label: string }> {
  return Object.fromEntries(
    actions.map(({ action, label }) => [action, { label }]),
  ) as Record<TKey, { label: string }>
}

function actionToolbarHandlers<TKey extends string>(
  actions: readonly ActionToolbarItem<TKey>[],
  onSelect: (action: TKey) => void,
): Partial<Record<TKey, () => void>> {
  return Object.fromEntries(
    actions.map(({ action }) => [action, () => onSelect(action)]),
  ) as Partial<Record<TKey, () => void>>
}

function omitAriaPressed(props: PatternButtonProps): PatternButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
