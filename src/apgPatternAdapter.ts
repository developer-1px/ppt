import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import type {
  PatternData,
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
      activeKey: activeKey ?? firstEnabledToolbarKey(rootKeys, disabledKeys),
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
  disabledKeys: readonly string[],
) {
  return keys.find((key) => !disabledKeys.includes(key)) ?? keys[0] ?? null
}

function omitAriaPressed(props: PatternButtonProps): PatternButtonProps {
  const next = { ...props }

  delete next['aria-pressed']

  return next
}
