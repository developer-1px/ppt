import {
  forwardRef,
  type ButtonHTMLAttributes,
  type HTMLAttributes,
} from 'react'
import { classNames } from './classNames'

type NativeButtonProps = ButtonHTMLAttributes<HTMLButtonElement>

export type ButtonProps = NativeButtonProps

export const Button = forwardRef<HTMLButtonElement, ButtonProps>(
  function Button({ className, type = 'button', ...props }, ref) {
    return (
      <button
        {...props}
        className={classNames('ppt-button', className)}
        ref={ref}
        type={type}
      />
    )
  },
)

export type IconButtonProps = Omit<
  NativeButtonProps,
  'aria-label' | 'title'
> & {
  label: string
  tooltip?: string
  variant?: 'floating' | 'icon' | 'slide-action'
}

export const IconButton = forwardRef<HTMLButtonElement, IconButtonProps>(
  function IconButton(
    {
      className,
      label,
      tooltip = label,
      type = 'button',
      variant = 'icon',
      ...props
    },
    ref,
  ) {
    return (
      <button
        {...props}
        aria-label={label}
        className={classNames(
          variant === 'floating'
            ? 'ppt-floating-command'
            : variant === 'slide-action'
              ? 'ppt-slide-action'
              : 'ppt-icon-button',
          className,
        )}
        ref={ref}
        title={tooltip}
        type={type}
      />
    )
  },
)

export type TabListProps = HTMLAttributes<HTMLDivElement>

export function TabList({ className, role = 'tablist', ...props }: TabListProps) {
  return (
    <div
      {...props}
      className={classNames('ppt-tab-list', className)}
      role={role}
    />
  )
}

export type TabProps = NativeButtonProps

export const Tab = forwardRef<HTMLButtonElement, TabProps>(
  function Tab({ className, type = 'button', ...props }, ref) {
    return (
      <button
        {...props}
        className={classNames('ppt-tab', className)}
        ref={ref}
        type={type}
      />
    )
  },
)

export type ToolbarGroupProps = HTMLAttributes<HTMLDivElement>

export function ToolbarGroup({
  className,
  role = 'group',
  ...props
}: ToolbarGroupProps) {
  return (
    <div
      {...props}
      className={classNames('ppt-toolbar-group', className)}
      role={role}
    />
  )
}

export type DisclosurePanelProps = HTMLAttributes<HTMLDivElement> & {
  open: boolean
}

export function DisclosurePanel({
  open,
  ...props
}: DisclosurePanelProps) {
  return <div {...props} data-ui-state={open ? 'open' : 'closed'} hidden={!open} />
}
