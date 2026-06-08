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

export function radioItemPropsByKey<TKey extends string>(
  items: readonly ReactRadioRenderItem[],
) {
  return Object.fromEntries(
    items.map((item) => [
      item.key,
      patternButtonProps(item.radioProps),
    ]),
  ) as Record<TKey, PatternButtonProps>
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
  return Object.fromEntries(
    tabs.map((tab) => [
      tab.value,
      patternButtonProps(tabRuntime.getTabProps(tab.tabKey)),
    ]),
  ) as Record<TValue, PatternButtonProps>
}

function omitAriaPressed(props: PatternButtonProps): PatternButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
