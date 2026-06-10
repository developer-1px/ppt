import {
  listboxDefinition,
  radioGroupDefinition,
  reducePatternData,
  tabsDefinition,
  toolbarDefinition,
} from '@interactive-os/aria/react'
import type {
  PatternData,
  PatternEvent,
} from '@interactive-os/aria/react'
import type {
  ManagedTabItem,
} from './apgPatternTypes'

export function radioGroupPatternData<TKey extends string>({
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

export function listboxPatternData<TKey extends string>({
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

export function tabsPatternData<TValue extends string>({
  activeValue,
  label,
  tabs,
}: {
  activeValue: TValue
  label: string
  tabs: readonly ManagedTabItem<TValue>[]
}): PatternData {
  const activeTab = activeManagedTab(activeValue, tabs)
  const controlsByKey: Record<string, readonly string[]> = {}
  const items: PatternData['items'] = {}
  const ownerByKey: Record<string, string> = {}
  const rootKeys: string[] = []

  for (const tab of tabs) {
    items[tab.tabKey] = { label: tab.label }
    items[tab.panelKey] = {
      label: tab.panelLabel ?? `${tab.label} view`,
    }
    controlsByKey[tab.tabKey] = [tab.panelKey]
    ownerByKey[tab.panelKey] = tab.tabKey
    rootKeys.push(tab.tabKey)
  }

  return {
    items,
    relations: {
      controlsByKey,
      ownerByKey,
      rootKeys,
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

function patternKey<TKey extends string>(key: string | null | undefined) {
  return key === null || key === undefined ? null : (key as TKey)
}

export function nextToolbarActiveKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
  disabledKeys: readonly TKey[] = [],
) {
  const nextKey = patternKey<TKey>(nextToolbarEventKey(data, event))

  if (nextKey === null || disabledKeys.includes(nextKey)) {
    return null
  }

  return nextKey
}

export function disabledToolbarKeys<TKey extends string>(
  disabledByKey: Partial<Record<TKey, boolean>>,
) {
  const disabledKeys: TKey[] = []

  for (const key in disabledByKey) {
    if (disabledByKey[key]) {
      disabledKeys.push(key)
    }
  }

  return disabledKeys
}

export function handleToolbarSelection<TKey extends string>(
  event: PatternEvent,
  handlers: Partial<Record<TKey, () => void>>,
) {
  for (const key of selectedPatternKeys(event)) {
    const handlerKey = patternKey<TKey>(key)

    if (handlerKey !== null) {
      handlers[handlerKey]?.()
    }
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

export function nextRadioActiveKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
): TKey | null {
  if (event.type === 'focus') {
    return patternKey<TKey>(event.key)
  }

  if (event.type === 'select') {
    return patternKey<TKey>(event.keys[0])
  }

  if (event.type === 'navigate') {
    return patternKey<TKey>(
      reducePatternData(
        radioGroupDefinition,
        data,
        event,
      ).state?.activeKey,
    )
  }

  return null
}

export function nextListboxSelectionKey<TKey extends string>(
  data: PatternData,
  event: PatternEvent,
): TKey | null {
  if (event.type === 'select') {
    return patternKey<TKey>(event.keys[0])
  }

  if (event.type === 'navigate') {
    return patternKey<TKey>(
      reducePatternData(
        listboxDefinition,
        data,
        event,
      ).state?.activeKey,
    )
  }

  return null
}

export function nextTabsSelectionValue<TValue extends string>(
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

export function activeManagedTab<TValue extends string>(
  activeValue: TValue,
  tabs: readonly ManagedTabItem<TValue>[],
) {
  const activeTab = tabs.find((tab) => tab.value === activeValue) ?? tabs[0]

  if (activeTab === undefined) {
    throw new Error('Managed tabs require at least one tab.')
  }

  return activeTab
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
  if (
    activeKey !== null &&
    activeKey !== undefined &&
    !disabledKeys.includes(activeKey)
  ) {
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

function tabValueFromKey<TValue extends string>(
  tabKey: string | null | undefined,
  tabs: readonly ManagedTabItem<TValue>[],
): TValue | null {
  return tabs.find((tab) => tab.tabKey === tabKey)?.value ?? null
}
