import type {
  ReactRadioRenderItem,
  ReactTabsRuntime,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'
import type {
  ManagedTabItem,
  PatternElementProps,
} from './apgPatternTypes'

export function toolbarItemPropsByKey(
  items: readonly ReactToolbarRenderItem[],
  options: { omitPressed?: boolean } = {},
) {
  const propsByKey: Record<string, PatternElementProps> = {}

  for (const item of items) {
    const props = item.itemProps

    if (options.omitPressed) {
      const nextProps = { ...props }

      delete nextProps['aria-pressed']
      propsByKey[item.key] = nextProps
      continue
    }

    propsByKey[item.key] = props
  }

  return propsByKey
}

export function radioItemPropsByKey(
  items: readonly ReactRadioRenderItem[],
) {
  const propsByKey: Record<string, PatternElementProps> = {}

  for (const item of items) {
    propsByKey[item.key] = item.radioProps
  }

  return propsByKey
}

export function tabsPropsByValue<TValue extends string>(
  tabRuntime: ReactTabsRuntime,
  tabs: readonly ManagedTabItem<TValue>[],
) {
  const propsByValue: Record<string, PatternElementProps> = {}

  for (const tab of tabs) {
    propsByValue[tab.value] = tabRuntime.getTabProps(tab.tabKey)
  }

  return propsByValue
}
