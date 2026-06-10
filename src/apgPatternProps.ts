import type {
  ReactListboxRenderItem,
  ReactRadioRenderItem,
  ReactTabsRuntime,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'
import type {
  ManagedTabItem,
  PatternButtonProps,
  PatternDivProps,
  PatternElementProps,
} from './apgPatternTypes'

function patternButtonProps(
  props: PatternElementProps,
): PatternButtonProps {
  return props as PatternButtonProps
}

export function patternDivProps(props: PatternElementProps): PatternDivProps {
  return props as PatternDivProps
}

export function toolbarItemPropsByKey(
  items: readonly ReactToolbarRenderItem[],
  options: { omitPressed?: boolean } = {},
) {
  const propsByKey: Record<string, PatternButtonProps> = {}

  for (const item of items) {
    const props = patternButtonProps(item.itemProps)

    propsByKey[item.key] = options.omitPressed
      ? omitAriaPressed(props)
      : props
  }

  return propsByKey
}

export function radioItemPropsByKey(
  items: readonly ReactRadioRenderItem[],
) {
  const propsByKey: Record<string, PatternButtonProps> = {}

  for (const item of items) {
    propsByKey[item.key] = patternButtonProps(item.radioProps)
  }

  return propsByKey
}

export function listboxRenderItems<TKey extends string>(
  items: readonly ReactListboxRenderItem[],
) {
  return items.map((item) => ({
    ...item,
    key: item.key as TKey,
    optionProps: patternButtonProps(item.optionProps),
  }))
}

export function tabsPropsByValue<TValue extends string>(
  tabRuntime: ReactTabsRuntime,
  tabs: readonly ManagedTabItem<TValue>[],
) {
  const propsByValue: Record<string, PatternButtonProps> = {}

  for (const tab of tabs) {
    propsByValue[tab.value] = patternButtonProps(
      tabRuntime.getTabProps(tab.tabKey),
    )
  }

  return propsByValue
}

function omitAriaPressed(props: PatternButtonProps): PatternButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
