import { X } from 'lucide-react'
import {
  getPPTCanvasModalBackdropPointerIntent,
  getPPTCanvasModalKeyboardIntent,
  PPT_MODAL_FOCUS_LIFECYCLE_MODEL,
  trapPPTCanvasModalTabFocus,
  usePPTCanvasModalFocusLifecycle,
} from '../../pptCanvasAppAffordanceAdapter'
import {
  useMemo,
  useRef,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from 'react'
import { IconButton } from '../core'
import type { PPTCommandPaletteItem } from './PPTSurfaceCommandModel'

type PPTShortcutHelpItem = {
  id: string
  section: string
  shortcut: string
  title: string
}

type PPTShortcutHelpSectionGroup = {
  items: PPTShortcutHelpItem[]
  section: string
}

type PPTShortcutHelpProps = {
  items: readonly PPTCommandPaletteItem[]
  onClose: () => void
  open: boolean
}

const PPT_SHORTCUT_HELP_SECTION_ORDER = [
  'Create',
  'Edit',
  'Arrange',
  'Format',
  'Slides',
  'View',
  'Export',
  'System',
]

export function PPTShortcutHelp({
  items,
  onClose,
  open,
}: PPTShortcutHelpProps) {
  if (!open) {
    return null
  }

  return <PPTShortcutHelpDialog items={items} onClose={onClose} />
}

function PPTShortcutHelpDialog({
  items,
  onClose,
}: Omit<PPTShortcutHelpProps, 'open'>) {
  const closeButtonRef = useRef<HTMLButtonElement | null>(null)
  const dialogRef = useRef<HTMLElement | null>(null)
  const groups = useMemo(
    () => groupPPTShortcutHelpItems(getPPTShortcutHelpItems(items)),
    [items],
  )

  usePPTCanvasModalFocusLifecycle({
    initialFocusRef: closeButtonRef,
  })

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    const backdropPointerIntent = getPPTCanvasModalBackdropPointerIntent({
      currentTarget: event.currentTarget,
      target: event.target,
    })

    if (backdropPointerIntent.kind === 'dismiss') {
      onClose()
    }
  }

  function handleKeyDown(event: ReactKeyboardEvent<HTMLElement>) {
    const modalKeyboardIntent = getPPTCanvasModalKeyboardIntent({
      key: event.key,
    })

    if (modalKeyboardIntent.kind === 'close') {
      if (modalKeyboardIntent.preventDefault) {
        event.preventDefault()
      }
      if (modalKeyboardIntent.stopPropagation) {
        event.stopPropagation()
      }
      onClose()
      return
    }

    if (modalKeyboardIntent.kind === 'trap-focus') {
      trapPPTCanvasModalTabFocus({
        event,
        root: dialogRef.current,
      })
    }
  }

  return (
    <div
      className="ppt-shortcut-help-backdrop"
      data-ppt-shortcut-help-backdrop
      onMouseDown={handleBackdropMouseDown}
    >
      <section
        aria-label="Keyboard shortcuts"
        aria-modal="true"
        className="ppt-shortcut-help"
        data-ppt-shortcut-help
        data-ppt-shortcut-help-focus-lifecycle={
          PPT_MODAL_FOCUS_LIFECYCLE_MODEL
        }
        ref={dialogRef}
        role="dialog"
        onKeyDown={handleKeyDown}
      >
        <header className="ppt-shortcut-help-header">
          <h2>Keyboard shortcuts</h2>
          <IconButton
            data-ppt-shortcut-help-close
            label="Close keyboard shortcuts"
            ref={closeButtonRef}
            tooltip="Close"
            onClick={onClose}
          >
            <X size={16} />
          </IconButton>
        </header>
        <div className="ppt-shortcut-help-sections">
          {groups.map((group) => (
            <section
              aria-label={group.section}
              className="ppt-shortcut-help-section"
              data-ppt-shortcut-help-section={group.section}
              key={group.section}
            >
              <h3>{group.section}</h3>
              <dl className="ppt-shortcut-help-list">
                {group.items.map((item) => (
                  <div
                    className="ppt-shortcut-help-row"
                    data-ppt-shortcut-help-item={item.id}
                    key={item.id}
                  >
                    <dt>{item.title}</dt>
                    <dd>
                      <kbd data-ppt-shortcut-help-shortcut={item.shortcut}>
                        {item.shortcut}
                      </kbd>
                    </dd>
                  </div>
                ))}
              </dl>
            </section>
          ))}
        </div>
      </section>
    </div>
  )
}

function getPPTShortcutHelpItems(
  items: readonly PPTCommandPaletteItem[],
): PPTShortcutHelpItem[] {
  return items.flatMap((item) =>
    item.shortcut
      ? [{
          id: item.id,
          section: item.section,
          shortcut: item.shortcut,
          title: item.title,
        }]
      : [],
  )
}

function groupPPTShortcutHelpItems(
  items: readonly PPTShortcutHelpItem[],
): PPTShortcutHelpSectionGroup[] {
  const orderedSections = [
    ...PPT_SHORTCUT_HELP_SECTION_ORDER,
    ...items
      .map((item) => item.section)
      .filter((section) => !PPT_SHORTCUT_HELP_SECTION_ORDER.includes(section)),
  ]

  return orderedSections.flatMap((section) => {
    const sectionItems = items.filter((item) => item.section === section)

    return sectionItems.length > 0
      ? [{ items: sectionItems, section }]
      : []
  })
}
