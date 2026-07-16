import { Circle, Diamond, Square } from 'lucide-react'
import {
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react'
import {
  PPT_MENU_ITEM_PROPS,
  PPT_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL,
  focusPPTCanvasElementOnNextFrame,
  getPPTCanvasMenuTriggerKeyboardIntent,
  usePPTCanvasMenuRovingFocus,
} from '../../pptCanvasAppAffordanceAdapter'
import type { PPTShapeKind } from '../../pptModel'
import { IconButton } from '../core'

export type PPTShapeQuickMenuState = {
  elementId: string
  shape: PPTShapeKind
}

export type PPTShapeKindMenuProps = {
  state: PPTShapeQuickMenuState
  onShapeKindChange: (elementId: string, shape: PPTShapeKind) => void
}

const SHAPE_OPTIONS = [{
  label: 'Rectangle',
  shape: 'rect',
}, {
  label: 'Oval',
  shape: 'ellipse',
}, {
  label: 'Diamond',
  shape: 'diamond',
}] as const satisfies readonly {
  label: string
  shape: PPTShapeKind
}[]

export function PPTShapeKindMenu({
  state,
  onShapeKindChange,
}: PPTShapeKindMenuProps) {
  const [open, setOpen] = useState(false)
  const [activeShape, setActiveShape] = useState<PPTShapeKind>(state.shape)
  const [initialActiveShapeIndex, setInitialActiveShapeIndex] = useState(0)
  const triggerRef = useRef<HTMLButtonElement | null>(null)
  const activeOption = SHAPE_OPTIONS.find((option) =>
    option.shape === activeShape
  ) ?? SHAPE_OPTIONS[0]
  const {
    onFocus: handleMenuFocus,
    onKeyDown: handleMenuKeyDown,
    ref: setMenuRoot,
  } = usePPTCanvasMenuRovingFocus<HTMLDivElement>({
    initialActiveIndex: initialActiveShapeIndex,
    onClose: () => {
      closeMenu()
      focusTrigger()
    },
  })

  function focusTrigger() {
    focusPPTCanvasElementOnNextFrame({
      resolveElement: () => triggerRef.current,
    })
  }

  function openMenu(shape = state.shape) {
    setInitialActiveShapeIndex(Math.max(
      0,
      SHAPE_OPTIONS.findIndex((option) => option.shape === shape),
    ))
    setOpen(true)
    setActiveShape(shape)
  }

  function closeMenu() {
    setOpen(false)
  }

  function commitShape(shape: PPTShapeKind) {
    onShapeKindChange(state.elementId, shape)
    closeMenu()
    focusTrigger()
  }

  function handleTriggerKeyDown(event: ReactKeyboardEvent<HTMLButtonElement>) {
    const intent = getPPTCanvasMenuTriggerKeyboardIntent({ key: event.key })

    if (intent.preventDefault) {
      event.preventDefault()
    }

    if (intent.kind === 'open-menu') {
      event.stopPropagation()
      openMenu()
    }
  }

  return (
    <span
      className="ppt-floating-menu-wrap"
      data-ppt-shape-menu-open={open ? 'true' : 'false'}
    >
      <IconButton
        aria-controls="ppt-shape-kind-menu"
        aria-expanded={open}
        aria-haspopup="menu"
        data-ppt-shape-menu-trigger
        label="Shape"
        ref={triggerRef}
        variant="floating"
        onClick={(event) => {
          event.preventDefault()
          event.stopPropagation()
          if (open) {
            closeMenu()
          } else {
            openMenu()
          }
        }}
        onKeyDown={handleTriggerKeyDown}
      >
        <PPTShapeIcon shape={state.shape} size={16} />
      </IconButton>
      {open ? (
        <div
          aria-label="Shape"
          className="ppt-floating-menu"
          data-ppt-shape-menu
          data-ppt-shape-menu-active={activeOption.shape}
          data-ppt-shape-menu-model={PPT_SELECTION_TOOLBAR_DROPDOWN_MENU_MODEL}
          id="ppt-shape-kind-menu"
          ref={setMenuRoot}
          role="menu"
          onFocus={handleMenuFocus}
          onKeyDown={handleMenuKeyDown}
        >
          {SHAPE_OPTIONS.map((option) => (
            <button
              {...PPT_MENU_ITEM_PROPS}
              aria-checked={state.shape === option.shape}
              className="ppt-floating-menu-item"
              data-ppt-shape-menu-item={option.shape}
              key={option.shape}
              role="menuitemcheckbox"
              tabIndex={option.shape === activeOption.shape ? 0 : -1}
              type="button"
              onClick={(event) => {
                event.preventDefault()
                event.stopPropagation()
                commitShape(option.shape)
              }}
              onFocus={() => setActiveShape(option.shape)}
              onMouseEnter={() => setActiveShape(option.shape)}
            >
              <PPTShapeIcon shape={option.shape} size={15} />
              <span>{option.label}</span>
            </button>
          ))}
        </div>
      ) : null}
    </span>
  )
}

function PPTShapeIcon({ shape, size }: { shape: PPTShapeKind; size: number }) {
  switch (shape) {
    case 'ellipse':
      return <Circle size={size} />
    case 'diamond':
      return <Diamond size={size} />
    case 'rect':
      return <Square size={size} />
  }
}
