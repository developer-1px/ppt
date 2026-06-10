import { useMemo, useState } from 'react'
import {
  useListboxPattern,
  useRadioGroupPattern,
  useTabsPattern,
  useToolbarPattern,
} from '@interactive-os/aria/react'
import {
  activeManagedTab,
  handleToolbarSelection,
  listboxPatternData,
  nextListboxSelectionKey,
  nextRadioActiveKey,
  nextTabsSelectionValue,
  nextToolbarActiveKey,
  radioGroupPatternData,
  toolbarPatternData,
  tabsPatternData,
} from './apgPatternData'
import {
  radioItemPropsByKey,
  tabsPropsByValue,
  toolbarItemPropsByKey,
} from './apgPatternProps'
import type {
  ActionToolbarItem,
  ListboxFocusStrategy,
  ManagedTabItem,
  ToolbarOrientation,
} from './apgPatternTypes'

export { disabledToolbarKeys } from './apgPatternData'

const EMPTY_TOOLBAR_KEYS: readonly never[] = []

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
  const data = useMemo(
    () =>
      tabsPatternData<TValue>({
        activeValue,
        label,
        tabs,
      }),
    [activeValue, label, tabs],
  )
  const tabRuntime = useTabsPattern(
    data,
    (event) => {
      const nextValue = nextTabsSelectionValue<TValue>(data, event, tabs)

      if (nextValue !== null) {
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
    panelProps: tabRuntime.getTabPanelProps(activeTab.panelKey),
    tablistProps: tabRuntime.getTablistProps(),
    tabPropsByValue: tabsPropsByValue(tabRuntime, tabs),
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
  const data = useMemo(
    () =>
      listboxPatternData<TKey>({
        activeKey,
        items,
        label,
        rootKeys,
      }),
    [activeKey, items, label, rootKeys],
  )
  const listbox = useListboxPattern(
    data,
    (event) => {
      const nextKey = nextListboxSelectionKey(data, event, rootKeys)

      if (nextKey !== null) {
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
    renderItems: listbox.renderItems,
    rootProps: listbox.rootProps,
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
  const data = useMemo(
    () =>
      radioGroupPatternData<TKey>({
        activeKey: activeRadioKey,
        items,
        label,
        rootKeys,
        selectedKey,
      }),
    [activeRadioKey, items, label, rootKeys, selectedKey],
  )
  const radio = useRadioGroupPattern(
    data,
    (event) => {
      const nextActiveKey = nextRadioActiveKey(data, event, rootKeys)

      if (nextActiveKey !== null) {
        setActiveRadioKey(nextActiveKey)
      }

      if (event.type === 'select' && nextActiveKey !== null) {
        onSelect(nextActiveKey)
      }
    },
    {
      elementIdPrefix,
    },
  )

  return {
    itemProps: radioItemPropsByKey(radio.renderItems),
    rootProps: radio.rootProps,
  }
}

function useManagedToolbarPattern<TKey extends string>({
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

  const data = useMemo(
    () =>
      toolbarPatternData<TKey>({
        activeKey: activeToolbarKey,
        disabledKeys,
        items,
        label,
        rootKeys,
        selectedKeys,
      }),
    [activeToolbarKey, disabledKeys, items, label, rootKeys, selectedKeys],
  )
  const toolbar = useToolbarPattern(
    data,
    (event) => {
      const nextActiveKey = nextToolbarActiveKey(
        data,
        event,
        rootKeys,
        disabledKeys,
      )

      if (nextActiveKey !== null) {
        setActiveToolbarKey(nextActiveKey)
      }

      handleToolbarSelection(event, handlers, rootKeys)
    },
    {
      elementIdPrefix,
      orientation,
    },
  )

  return {
    itemProps: toolbarItemPropsByKey(toolbar.renderItems, {
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
  const toolbarData = useMemo(() => {
    const handlers: Partial<Record<TKey, () => void>> = {}
    const items: Record<string, { label: string }> = {}
    const rootKeys: TKey[] = []

    for (const { action, label } of actions) {
      handlers[action] = () => onSelect(action)
      items[action] = { label }
      rootKeys.push(action)
    }

    return {
      handlers,
      items,
      rootKeys,
    }
  }, [actions, onSelect])

  return useManagedToolbarPattern<TKey>({
    activeKey,
    disabledKeys,
    elementIdPrefix,
    handlers: toolbarData.handlers,
    items: toolbarData.items,
    label,
    omitPressed,
    orientation,
    rootKeys: toolbarData.rootKeys,
    selectedKeys,
  })
}
