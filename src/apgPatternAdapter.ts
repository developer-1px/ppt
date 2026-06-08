import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import type {
  PatternEvent,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'

type PatternButtonProps = ButtonHTMLAttributes<HTMLButtonElement>
type PatternDivProps = HTMLAttributes<HTMLDivElement>
type PatternElementProps = HTMLAttributes<HTMLElement>

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
  disabledKeys: readonly string[],
) {
  return keys.find((key) => !disabledKeys.includes(key)) ?? keys[0] ?? null
}

function omitAriaPressed(props: PatternButtonProps): PatternButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
