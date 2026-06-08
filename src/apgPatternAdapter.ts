import { useMemo, useState } from 'react'
import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import {
  listboxDefinition,
  radioGroupDefinition,
  reducePatternData,
  tabsDefinition,
  toolbarDefinition,
  useListboxPattern,
  useRadioGroupPattern,
  useTabsPattern,
  useToolbarPattern,
} from '@interactive-os/aria/react'
import type {
  PatternData,
  PatternEvent,
  PatternOptions,
  ReactListboxRenderItem,
  ReactRadioRenderItem,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'

type PatternButtonProps = ButtonHTMLAttributes<HTMLButtonElement>
type PatternDivProps = HTMLAttributes<HTMLDivElement>
type PatternElementProps = HTMLAttributes<HTMLElement>
type ToolbarOrientation = NonNullable<PatternOptions['orientation']>
type ListboxFocusStrategy = NonNullable<PatternOptions['focusStrategy']>

const EMPTY_TOOLBAR_KEYS: readonly never[] = []

type ManagedTabItem<TValue extends string> = {
  label: string
  panelKey: string
  panelLabel?: string
  tabKey: string
  value: TValue
}

export type ActionToolbarItem<TKey extends string> = {
  action: TKey
  label: string
}

export function useManagedTabsPattern<TValue extends string>({
  activeValue,
  elementIdPrefix,
  label,
  onSelect,
  tabs,
}: {
  activeValue: TValue
  elementIdPrefix: string
  label: string
  onSelect: (value: TValue) => void
  tabs: readonly ManagedTabItem<TValue>[]
}) {
  const data = useMemo(() =>
    tabsPatternData<TValue>({
      activeValue,
      label,
      tabs,
    }), [activeValue, label, tabs])
  const tabRuntime = useTabsPattern(
    data,
    (event) => {
      const nextValue = nextTabsSelectionValue<TValue>(data, event, tabs)

      if (nextValue) {
        onSelect(nextValue)
      }
    },
    {
      activationMode: 'automatic',
      elementIdPrefix,
      orientation: 'horizontal',
    },
  )
  const activeTab = activeManagedTab(activeValue, tabs)

  return {
    panelProps: patternDivProps(tabRuntime.getTabPanelProps(activeTab.panelKey)),
    tablistProps: patternDivProps(tabRuntime.getTablistProps()),
    tabPropsByValue: tabsPropsByValue<TValue>(tabRuntime, tabs),
  }
}

export function useManagedListboxPattern<TKey extends string>({
  activeKey,
  elementIdPrefix,
  focusStrategy = 'rovingTabIndex',
  items,
  label,
  onSelect,
  rootKeys,
  typeaheadEnabled = true,
}: {
  activeKey: TKey
  elementIdPrefix: string
  focusStrategy?: ListboxFocusStrategy
  items: Record<TKey, { label: string; textValue?: string }>
  label: string
  onSelect: (key: TKey) => void
  rootKeys: readonly TKey[]
  typeaheadEnabled?: boolean
}) {
  const data = useMemo(() =>
    listboxPatternData<TKey>({
      activeKey,
      items,
      label,
      rootKeys,
    }), [activeKey, items, label, rootKeys])
  const listbox = useListboxPattern(
    data,
    (event) => {
      const nextKey = nextListboxSelectionKey<TKey>(data, event)

      if (nextKey) {
        onSelect(nextKey)
      }
    },
    {
      elementIdPrefix,
      focusStrategy,
      orientation: 'vertical',
      selectionMode: 'single',
      typeaheadEnabled,
    },
  )

  return {
    renderItems: listboxRenderItems<TKey>(listbox.renderItems),
    rootProps: patternDivProps(listbox.rootProps),
  }
}

export function useManagedRadioGroupPattern<TKey extends string>({
  elementIdPrefix,
  items,
  label,
  onSelect,
  rootKeys,
  selectedKey,
}: {
  elementIdPrefix: string
  items: Record<TKey, { label: string }>
  label: string
  onSelect: (key: TKey) => void
  rootKeys: readonly TKey[]
  selectedKey: TKey
}) {
  const [activeRadioKey, setActiveRadioKey] = useState<TKey | null>(selectedKey)
  const data = useMemo(() =>
    radioGroupPatternData<TKey>({
      activeKey: activeRadioKey,
      items,
      label,
      rootKeys,
      selectedKey,
    }), [activeRadioKey, items, label, rootKeys, selectedKey])
  const radio = useRadioGroupPattern(
    data,
    (event) => {
      const nextActiveKey = nextRadioActiveKey<TKey>(data, event)

      if (nextActiveKey) {
        setActiveRadioKey(nextActiveKey)
      }

      for (const key of selectedPatternKeys(event)) {
        onSelect(key as TKey)
      }
    },
    {
      elementIdPrefix,
    },
  )

  return {
    itemProps: radioItemPropsByKey<TKey>(radio.renderItems),
    rootProps: patternDivProps(radio.rootProps),
  }
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

function radioGroupPatternData<TKey extends string>({
  activeKey,
  items,
  label,
  rootKeys,
  selectedKey,
}: {
  activeKey?: TKey | null
  items: Record<TKey, { label: string }>
  label: string
  rootKeys: readonly TKey[]
  selectedKey: TKey
}): PatternData {
  return {
    items,
    relations: { rootKeys },
    refs: { label },
    state: {
      activeKey: activeKey ?? selectedKey,
      selectedKeys: [selectedKey],
    },
  }
}

function listboxPatternData<TKey extends string>({
  activeKey,
  items,
  label,
  rootKeys,
}: {
  activeKey: TKey
  items: Record<TKey, { label: string; textValue?: string }>
  label: string
  rootKeys: readonly TKey[]
}): PatternData {
  return {
    items,
    relations: { rootKeys },
    refs: { label },
    state: {
      activeKey,
      selectedKeys: [activeKey],
    },
  }
}

function tabsPatternData<TValue extends string>({
  activeValue,
  label,
  tabs,
}: {
  activeValue: TValue
  label: string
  tabs: readonly ManagedTabItem<TValue>[]
}): PatternData {
  const activeTab = activeManagedTab(activeValue, tabs)

  return {
    items: Object.fromEntries(
      tabs.flatMap((tab) => [
        [tab.tabKey, { label: tab.label }],
        [tab.panelKey, { label: tab.panelLabel ?? `${tab.label} view` }],
      ]),
    ),
    relations: {
      controlsByKey: Object.fromEntries(
        tabs.map((tab) => [tab.tabKey, [tab.panelKey]]),
      ),
      ownerByKey: Object.fromEntries(
        tabs.map((tab) => [tab.panelKey, tab.tabKey]),
      ),
      rootKeys: tabs.map((tab) => tab.tabKey),
    },
    refs: { label },
    state: {
      activeKey: activeTab.tabKey,
      selectedKeys: [activeTab.tabKey],
    },
  }
}

function selectedPatternKeys(event: PatternEvent) {
  return event.type === 'select' ? event.keys : []
}

function nextToolbarActiveKey<TKey extends string>(
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

function handleToolbarSelection<TKey extends string>(
  event: PatternEvent,
  handlers: Partial<Record<TKey, () => void>>,
) {
  for (const key of selectedPatternKeys(event)) {
    handlers[key as TKey]?.()
  }
}

function toolbarPatternData<TKey extends string>({
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

function toolbarItemPropsByKey<TKey extends string>(
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

function radioItemPropsByKey<TKey extends string>(
  items: readonly ReactRadioRenderItem[],
) {
  return Object.fromEntries(
    items.map((item) => [
      item.key,
      patternButtonProps(item.radioProps),
    ]),
  ) as Record<TKey, PatternButtonProps>
}

function listboxRenderItems<TKey extends string>(
  items: readonly ReactListboxRenderItem[],
) {
  return items.map((item) => ({
    ...item,
    key: item.key as TKey,
    optionProps: patternButtonProps(item.optionProps),
  }))
}

function tabsPropsByValue<TValue extends string>(
  tabRuntime: ReturnType<typeof useTabsPattern>,
  tabs: readonly ManagedTabItem<TValue>[],
) {
  return Object.fromEntries(
    tabs.map((tab) => [
      tab.value,
      patternButtonProps(tabRuntime.getTabProps(tab.tabKey)),
    ]),
  ) as Record<TValue, PatternButtonProps>
}

function firstEnabledToolbarKey<TKey extends string>(
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

  return selectedPatternKeys(event)[0] ?? null
}

function nextRadioActiveKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
): TKey | null {
  if (event.type === 'focus') {
    return event.key as TKey
  }

  if (event.type === 'select') {
    return event.keys[0] as TKey | undefined ?? null
  }

  if (event.type === 'navigate') {
    return reducePatternData(
      radioGroupDefinition,
      data,
      event,
    ).state?.activeKey as TKey | undefined ?? null
  }

  return null
}

function nextListboxSelectionKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
): TKey | null {
  if (event.type === 'select') {
    return event.keys[0] as TKey | undefined ?? null
  }

  if (event.type === 'navigate') {
    return reducePatternData(
      listboxDefinition,
      data,
      event,
    ).state?.activeKey as TKey | undefined ?? null
  }

  return null
}

function nextTabsSelectionValue<TValue extends string>(
  data: PatternData,
  event: PatternEvent,
  tabs: readonly ManagedTabItem<TValue>[],
): TValue | null {
  if (event.type === 'select') {
    return tabValueFromKey(event.keys[0], tabs)
  }

  if (event.type === 'navigate') {
    return tabValueFromKey(
      reducePatternData(tabsDefinition, data, event).state?.activeKey,
      tabs,
    )
  }

  return null
}

function tabValueFromKey<TValue extends string>(
  tabKey: string | null | undefined,
  tabs: readonly ManagedTabItem<TValue>[],
): TValue | null {
  return tabs.find((tab) => tab.tabKey === tabKey)?.value ?? null
}

function activeManagedTab<TValue extends string>(
  activeValue: TValue,
  tabs: readonly ManagedTabItem<TValue>[],
) {
  const activeTab = tabs.find((tab) => tab.value === activeValue) ?? tabs[0]

  if (!activeTab) {
    throw new Error('Managed tabs require at least one tab.')
  }

  return activeTab
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
