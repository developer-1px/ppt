import { useEffect, type CSSProperties } from 'react'
import { ChevronLeft, ChevronRight, X } from 'lucide-react'
import { rectToStyle, type RetouchSlide } from './retouchModel'

type PresentationOverlayProps = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  notes: string
  onClose: () => void
  onNext: () => void
  onPrevious: () => void
  slideCount: number
}

export function PresentationOverlay({
  activeSlide,
  activeSlideIndex,
  notes,
  onClose,
  onNext,
  onPrevious,
  slideCount,
}: PresentationOverlayProps) {
  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent) {
      if (event.key === 'Escape') {
        event.preventDefault()
        onClose()
        return
      }

      if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
        event.preventDefault()
        onNext()
        return
      }

      if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
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
            data-slide={activeSlide.id}
            style={{ '--accent': activeSlide.accent } as CSSProperties}
          >
            {activeSlide.blocks.map((block) => (
              <div
                className={`presentation-block ${block.className}`}
                data-block={block.id}
                data-role={block.role}
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

      <div className="presentation-controls" role="toolbar" aria-label="Presentation">
        <button
          aria-label="Previous slide"
          disabled={activeSlideIndex === 0}
          onClick={onPrevious}
          type="button"
        >
          <ChevronLeft aria-hidden="true" size={18} />
        </button>
        <button aria-label="Close presentation" onClick={onClose} type="button">
          <X aria-hidden="true" size={18} />
        </button>
        <button
          aria-label="Next slide"
          disabled={activeSlideIndex === slideCount - 1}
          onClick={onNext}
          type="button"
        >
          <ChevronRight aria-hidden="true" size={18} />
        </button>
      </div>
    </div>
  )
}
