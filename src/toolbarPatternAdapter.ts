import type { ButtonHTMLAttributes } from 'react'
import type {
  PatternEvent,
  ReactToolbarRenderItem,
} from '@interactive-os/aria/react'

type ToolbarButtonProps = ButtonHTMLAttributes<HTMLButtonElement>

export function selectedToolbarKeys(event: PatternEvent) {
  return event.type === 'select' ? event.keys : []
}

export function toolbarItemPropsByKey<TKey extends string>(
  items: readonly ReactToolbarRenderItem[],
  options: { omitPressed?: boolean } = {},
) {
  return Object.fromEntries(
    items.map((item) => {
      const props = item.itemProps as ToolbarButtonProps

      return [
        item.key,
        options.omitPressed ? omitAriaPressed(props) : props,
      ]
    }),
  ) as Record<TKey, ToolbarButtonProps>
}

export function firstEnabledToolbarKey<TKey extends string>(
  keys: readonly TKey[],
  disabledKeys: readonly string[],
) {
  return keys.find((key) => !disabledKeys.includes(key)) ?? keys[0] ?? null
}

function omitAriaPressed(props: ToolbarButtonProps): ToolbarButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
