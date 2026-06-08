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

type PresentationOverlayProps = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  notes: string
  onClose: () => void
  onNext: () => void
  onPrevious: () => void
  slideCount: number
}

type PresentationControlKey = 'close' | 'next' | 'previous'

const PRESENTATION_CONTROL_ACTIONS: {
  action: PresentationControlKey
  label: string
}[] = [
  { action: 'previous', label: 'Previous slide' },
  { action: 'close', label: 'Close presentation' },
  { action: 'next', label: 'Next slide' },
]

export function PresentationOverlay({
  activeSlide,
  activeSlideIndex,
  notes,
  onClose,
  onNext,
  onPrevious,
  slideCount,
}: PresentationOverlayProps) {
  const controlToolbar = useActionToolbarPattern<PresentationControlKey>({
    actions: PRESENTATION_CONTROL_ACTIONS,
    activeKey: 'close',
    disabledKeys: disabledToolbarKeys<PresentationControlKey>([
      ['previous', activeSlideIndex === 0],
      ['next', activeSlideIndex === slideCount - 1],
    ]),
    elementIdPrefix: 'presentation-control-',
    label: 'Presentation',
    onSelect: (action) => {
      if (action === 'close') onClose()
      if (action === 'next') onNext()
      if (action === 'previous') onPrevious()
    },
  })

  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent) {
      if (matchesShortcut(event, 'Escape')) {
        event.preventDefault()
        onClose()
        return
      }

      if (isPresentationControlTarget(event.target)) {
        return
      }

      if (matchesShortcut(event, 'ArrowRight PageDown Space')) {
        event.preventDefault()
        onNext()
        return
      }

      if (matchesShortcut(event, 'ArrowLeft PageUp')) {
        event.preventDefault()
        onPrevious()
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
        <button
          {...controlToolbar.itemProps.previous}
          aria-label="Previous slide"
          disabled={activeSlideIndex === 0}
          type="button"
        >
          <ChevronLeft aria-hidden="true" size={18} />
        </button>
        <button
          {...controlToolbar.itemProps.close}
          aria-label="Close presentation"
          type="button"
        >
          <X aria-hidden="true" size={18} />
        </button>
        <button
          {...controlToolbar.itemProps.next}
          aria-label="Next slide"
          disabled={activeSlideIndex === slideCount - 1}
          type="button"
        >
          <ChevronRight aria-hidden="true" size={18} />
        </button>
      </div>
    </div>
  )
}

function isPresentationControlTarget(target: EventTarget | null) {
  return target instanceof Element && Boolean(target.closest('.presentation-controls'))
}
