import { useEffect, useMemo, useState } from 'react'
import {
  reducePatternData,
  toolbarDefinition,
  useToolbarPattern,
} from '@interactive-os/aria/react'
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
  handleToolbarSelection,
  toolbarPatternData,
  toolbarItemPropsByKey,
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

const PRESENTATION_CONTROL_KEYS = ['previous', 'close', 'next'] as const

export function PresentationOverlay({
  activeSlide,
  activeSlideIndex,
  notes,
  onClose,
  onNext,
  onPrevious,
  slideCount,
}: PresentationOverlayProps) {
  const [activeControlKey, setActiveControlKey] = useState<PresentationControlKey>('close')
  const controlDisabledKeys = useMemo(() =>
    disabledToolbarKeys<PresentationControlKey>([
      ['previous', activeSlideIndex === 0],
      ['next', activeSlideIndex === slideCount - 1],
    ]), [activeSlideIndex, slideCount])
  const controlToolbarData = useMemo(() =>
    toolbarPatternData<PresentationControlKey>({
      disabledKeys: controlDisabledKeys,
      items: {
        close: { label: 'Close presentation' },
        next: { label: 'Next slide' },
        previous: { label: 'Previous slide' },
      },
      label: 'Presentation',
      rootKeys: PRESENTATION_CONTROL_KEYS,
      activeKey: controlDisabledKeys.includes(activeControlKey)
        ? null
        : activeControlKey,
    }), [activeControlKey, controlDisabledKeys])
  const controlToolbar = useToolbarPattern(
    controlToolbarData,
    (event) => {
      if (event.type === 'navigate') {
        const nextActiveKey = reducePatternData(
          toolbarDefinition,
          controlToolbarData,
          event,
        ).state?.activeKey as PresentationControlKey | undefined

        if (nextActiveKey) {
          setActiveControlKey(nextActiveKey)
        }
      }

      handleToolbarSelection<PresentationControlKey>(event, {
        close: onClose,
        next: onNext,
        previous: onPrevious,
      })
    },
    {
      elementIdPrefix: 'presentation-control-',
      orientation: 'horizontal',
    },
  )
  const controlProps = toolbarItemPropsByKey<PresentationControlKey>(
    controlToolbar.renderItems,
    { omitPressed: true },
  )

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
          {...controlProps.previous}
          aria-label="Previous slide"
          disabled={activeSlideIndex === 0}
          type="button"
        >
          <ChevronLeft aria-hidden="true" size={18} />
        </button>
        <button
          {...controlProps.close}
          aria-label="Close presentation"
          type="button"
        >
          <X aria-hidden="true" size={18} />
        </button>
        <button
          {...controlProps.next}
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
