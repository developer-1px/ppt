import { type CSSProperties } from 'react'
import {
  rectToStyle,
  type RetouchSlide,
} from './retouchModel'

type DeckGridProps = {
  activeSlideId: string
  changedSlideIds: Set<string>
  onOpenSlide: (slideId: string) => void
  slides: RetouchSlide[]
}

export function DeckGrid({
  activeSlideId,
  changedSlideIds,
  onOpenSlide,
  slides,
}: DeckGridProps) {
  return (
    <div className="deck-grid" aria-label="Deck grid">
      {slides.map((slide, index) => (
        <button
          aria-current={slide.id === activeSlideId ? 'page' : undefined}
          className="grid-slide-card"
          data-changed={changedSlideIds.has(slide.id) ? 'true' : 'false'}
          key={slide.id}
          onClick={() => onOpenSlide(slide.id)}
          type="button"
        >
          <span className="grid-slide-index">{index + 1}</span>
          <span
            className="grid-slide-canvas"
            data-slide={slide.id}
            style={{ '--accent': slide.accent } as CSSProperties}
          >
            {slide.blocks.map((block) => (
              <span
                className={`grid-slide-block ${block.className}`}
                data-role={block.role}
                key={block.id}
                style={rectToStyle(block)}
              >
                {block.text}
              </span>
            ))}
          </span>
          <span className="grid-slide-name">{slide.name}</span>
        </button>
      ))}
    </div>
  )
}
