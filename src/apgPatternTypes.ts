import type { ButtonHTMLAttributes, HTMLAttributes } from 'react'
import type { PatternOptions } from '@interactive-os/aria/react'

export type PatternButtonProps = ButtonHTMLAttributes<HTMLButtonElement>
export type PatternDivProps = HTMLAttributes<HTMLDivElement>
export type PatternElementProps = HTMLAttributes<HTMLElement>
export type ToolbarOrientation = NonNullable<PatternOptions['orientation']>
export type ListboxFocusStrategy = NonNullable<PatternOptions['focusStrategy']>

export type ManagedTabItem<TValue extends string> = {
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
