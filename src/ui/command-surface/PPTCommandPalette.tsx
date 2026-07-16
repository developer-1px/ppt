import {
  filterPPTCommandPaletteItems,
  getPPTCanvasCommandPaletteKeyboardIntent,
  getPPTCanvasModalBackdropPointerIntent,
  getPPTCanvasModalKeyboardIntent,
  PPT_COMMAND_PALETTE_ITEMS_MODEL,
  PPT_MODAL_FOCUS_LIFECYCLE_MODEL,
  trapPPTCanvasModalTabFocus,
  usePPTCanvasModalFocusLifecycle,
} from '../../pptCanvasAppAffordanceAdapter'
import {
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type MouseEvent as ReactMouseEvent,
} from 'react'
import type { PPTCommandPaletteItem } from './PPTSurfaceCommandModel'

type PPTCommandPaletteProps = {
  items: readonly PPTCommandPaletteItem[]
  onClose: () => void
  open: boolean
}

export function PPTCommandPalette({
  items,
  onClose,
  open,
}: PPTCommandPaletteProps) {
  if (!open) {
    return null
  }

  return <PPTCommandPaletteDialog items={items} onClose={onClose} />
}

function PPTCommandPaletteDialog({
  items,
  onClose,
}: Omit<PPTCommandPaletteProps, 'open'>) {
  const [query, setQuery] = useState('')
  const [activeIndex, setActiveIndex] = useState(0)
  const dialogRef = useRef<HTMLElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const filteredItems = useMemo(
    () => filterPPTCommandPaletteItems(items, query).slice(0, 10),
    [items, query],
  )
  const maxActiveIndex = Math.max(0, filteredItems.length - 1)
  const activeItemIndex = Math.min(activeIndex, maxActiveIndex)
  const activeItem = filteredItems[activeItemIndex]
  const listboxId = 'ppt-command-palette-listbox'
  const activeOptionId = activeItem
    ? getPPTCommandPaletteOptionId(activeItem.id)
    : undefined

  usePPTCanvasModalFocusLifecycle({
    initialFocusRef: inputRef,
  })

  function runItem(item: PPTCommandPaletteItem | undefined) {
    if (!item || item.disabled) {
      return
    }

    item.onSelect()
    onClose()
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
      return
    }

    const keyboardIntent = getPPTCanvasCommandPaletteKeyboardIntent({
      activeIndex: activeItemIndex,
      itemCount: filteredItems.length,
      key: event.key,
    })

    if (!keyboardIntent.preventDefault) {
      return
    }

    event.preventDefault()
    event.stopPropagation()

    if (keyboardIntent.kind === 'move-active') {
      setActiveIndex(keyboardIntent.activeIndex)
      return
    }

    if (keyboardIntent.kind === 'run-active') {
      runItem(filteredItems[keyboardIntent.activeIndex])
    }
  }

  function handleBackdropMouseDown(event: ReactMouseEvent<HTMLDivElement>) {
    const backdropPointerIntent = getPPTCanvasModalBackdropPointerIntent({
      currentTarget: event.currentTarget,
      target: event.target,
    })

    if (backdropPointerIntent.kind === 'dismiss') {
      onClose()
    }
  }

  return (
    <div
      className="ppt-command-palette-backdrop"
      data-ppt-command-palette-backdrop
      onMouseDown={handleBackdropMouseDown}
    >
      <section
        aria-label="Command palette"
        aria-modal="true"
        className="ppt-command-palette"
        data-ppt-command-palette
        data-ppt-command-palette-focus-lifecycle={
          PPT_MODAL_FOCUS_LIFECYCLE_MODEL
        }
        data-ppt-command-palette-focus-trap="true"
        data-ppt-command-palette-model={PPT_COMMAND_PALETTE_ITEMS_MODEL}
        data-ppt-command-palette-restore-focus="true"
        ref={dialogRef}
        role="dialog"
        onKeyDown={handleKeyDown}
      >
        <input
          aria-activedescendant={activeOptionId}
          aria-autocomplete="list"
          aria-controls={listboxId}
          aria-expanded="true"
          aria-label="Search commands"
          className="ppt-command-palette-input"
          data-ppt-command-palette-active-option={activeOptionId}
          data-ppt-command-palette-combobox="true"
          data-ppt-command-palette-controls={listboxId}
          data-ppt-command-palette-query
          placeholder="Search commands"
          ref={inputRef}
          role="combobox"
          value={query}
          onChange={(event) => {
            setQuery(event.target.value)
            setActiveIndex(0)
          }}
        />
        <div
          aria-label="Command results"
          className="ppt-command-palette-list"
          data-ppt-command-palette-active-option={activeOptionId}
          data-ppt-command-palette-listbox
          id={listboxId}
          role="listbox"
        >
          {filteredItems.length > 0 ? (
            filteredItems.map((item, index) => (
              <button
                aria-disabled={item.disabled ? 'true' : undefined}
                aria-selected={index === activeItemIndex}
                className="ppt-command-palette-item"
                data-ppt-command-palette-active={
                  index === activeItemIndex ? 'true' : undefined
                }
                data-ppt-command-palette-item={item.id}
                data-ppt-command-palette-option-id={
                  getPPTCommandPaletteOptionId(item.id)
                }
                disabled={item.disabled}
                id={getPPTCommandPaletteOptionId(item.id)}
                key={item.id}
                role="option"
                type="button"
                onClick={() => runItem(item)}
                onMouseEnter={() => setActiveIndex(index)}
              >
                <span className="ppt-command-palette-item-main">
                  <span className="ppt-command-palette-item-title">
                    {item.title}
                  </span>
                  <span className="ppt-command-palette-item-section">
                    {item.section}
                  </span>
                </span>
                {item.shortcut ? (
                  <kbd className="ppt-command-palette-shortcut">
                    {item.shortcut}
                  </kbd>
                ) : null}
              </button>
            ))
          ) : (
            <div
              className="ppt-command-palette-empty"
              data-ppt-command-palette-empty
            >
              No matches
            </div>
          )}
        </div>
      </section>
    </div>
  )
}

function getPPTCommandPaletteOptionId(itemId: string) {
  return `ppt-command-palette-option-${itemId.replace(/[^A-Za-z0-9_-]/g, '-')}`
}
