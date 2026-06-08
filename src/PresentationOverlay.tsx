import { useEffect } from 'react'
import { ChevronLeft, ChevronRight, X } from 'lucide-react'
import { matchesShortcut } from '@interactive-os/keyboard'
import { rectToStyle, type RetouchSlide } from './retouchModel'
import {
  htmlSlideBlockAttributes,
  htmlSlideRootAttributes,
} from './htmlSlideContract'
import { cssVariables } from './cssVariables'
import {
  disabledToolbarKeys,
  useActionToolbarPattern,
} from './apgPatternAdapter'
import './PresentationOverlay.css'

type PresentationOverlayProps = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  notes: string
  onClose: () => void
  onNext: () => void
  onPrevious: () => void
  slideCount: number
}

const PRESENTATION_CONTROL_ACTIONS = [
  { action: 'previous', icon: ChevronLeft, label: 'Previous slide' },
  { action: 'close', icon: X, label: 'Close presentation' },
  { action: 'next', icon: ChevronRight, label: 'Next slide' },
] as const

type PresentationControlKey =
  (typeof PRESENTATION_CONTROL_ACTIONS)[number]['action']

export function PresentationOverlay({
  activeSlide,
  activeSlideIndex,
  notes,
  onClose,
  onNext,
  onPrevious,
  slideCount,
}: PresentationOverlayProps) {
  const presentationControlCommands = {
    close: onClose,
    next: onNext,
    previous: onPrevious,
  } satisfies Record<PresentationControlKey, () => void>
  const presentationControlDisabled = {
    close: false,
    next: activeSlideIndex === slideCount - 1,
    previous: activeSlideIndex === 0,
  } satisfies Record<PresentationControlKey, boolean>

  const controlToolbar = useActionToolbarPattern<PresentationControlKey>({
    actions: PRESENTATION_CONTROL_ACTIONS,
    activeKey: 'close',
    disabledKeys: disabledToolbarKeys<PresentationControlKey>(
      presentationControlDisabled,
    ),
    elementIdPrefix: 'presentation-control-',
    label: 'Presentation',
    onSelect: (action) => presentationControlCommands[action](),
  })

  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent) {
      if (matchesShortcut(event, 'Escape')) {
        consumePresentationShortcut(event, onClose)
        return
      }

      if (isPresentationControlTarget(event.target)) {
        return
      }

      if (matchesShortcut(event, 'ArrowRight PageDown Space')) {
        consumePresentationShortcut(event, onNext)
        return
      }

      if (matchesShortcut(event, 'ArrowLeft PageUp')) {
        consumePresentationShortcut(event, onPrevious)
      }
    }

    window.addEventListener('keydown', handleKeyDown)

    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [onClose, onNext, onPrevious])

  return (
    <div
      aria-label="Presentation preview"
      className="presentation-overlay"
      role="dialog"
      aria-modal="true"
    >
      <div className="presentation-stage">
        <div className="presentation-slide-frame">
          <div
            className="presentation-slide"
            {...htmlSlideRootAttributes(activeSlide.id)}
            style={cssVariables({ '--accent': activeSlide.accent })}
          >
            {activeSlide.blocks.map((block, blockIndex) => (
              <div
                className={`presentation-block ${block.className}`}
                {...htmlSlideBlockAttributes(block, blockIndex)}
                key={block.id}
                style={rectToStyle(block)}
              >
                {block.text}
              </div>
            ))}
          </div>
        </div>
      </div>

      <aside className="presentation-notes" aria-label="Presenter notes">
        <div className="presentation-meta">
          <span>{activeSlideIndex + 1} / {slideCount}</span>
          <strong>{activeSlide.name}</strong>
        </div>
        <textarea aria-label="Notes" readOnly value={notes} />
      </aside>

      <div {...controlToolbar.rootProps} className="presentation-controls">
        {PRESENTATION_CONTROL_ACTIONS.map(({ action, icon: Icon, label }) => (
          <button
            {...controlToolbar.itemProps[action]}
            aria-label={label}
            disabled={presentationControlDisabled[action]}
            key={action}
            type="button"
          >
            <Icon aria-hidden="true" size={18} />
          </button>
        ))}
      </div>
    </div>
  )
}

function consumePresentationShortcut(event: KeyboardEvent, action: () => void) {
  event.preventDefault()
  action()
}

function isPresentationControlTarget(target: EventTarget | null) {
  return target instanceof Element && Boolean(target.closest('.presentation-controls'))
}
