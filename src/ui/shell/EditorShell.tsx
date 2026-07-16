import {
  forwardRef,
  type HTMLAttributes,
  type ReactNode,
} from 'react'
import {
  classNames,
  type EditorTransientSurface,
} from '../core'

export type EditorTheme = 'dark' | 'light'

export type EditorShellProps = HTMLAttributes<HTMLElement> & {
  inspectorOpen: boolean
  theme: EditorTheme
}

export const EditorShell = forwardRef<HTMLElement, EditorShellProps>(
  function EditorShell(
    { children, className, inspectorOpen, theme, ...props },
    ref,
  ) {
    return (
      <main
        {...props}
        className={classNames('ppt-app', className)}
        data-editor-inspector-open={inspectorOpen ? 'true' : 'false'}
        data-editor-shell
        data-theme={theme}
        ref={ref}
      >
        {children}
      </main>
    )
  },
)

export type EditorToolbarProps = HTMLAttributes<HTMLElement> & {
  activeTransientSurface: EditorTransientSurface | null
  children: ReactNode
}

export const EditorToolbar = forwardRef<HTMLElement, EditorToolbarProps>(
  function EditorToolbar(
    {
      activeTransientSurface,
      children,
      className,
      role = 'toolbar',
      ...props
    },
    ref,
  ) {
    return (
      <header
        {...props}
        aria-orientation="horizontal"
        className={classNames('ppt-topbar', className)}
        data-editor-toolbar
        data-editor-transient-surface={activeTransientSurface ?? undefined}
        ref={ref}
        role={role}
      >
        {children}
      </header>
    )
  },
)
