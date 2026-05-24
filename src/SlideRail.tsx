import { type ButtonHTMLAttributes, type CSSProperties, useMemo } from 'react'
import {
  ArrowDown,
  ArrowUp,
  Copy,
  Plus,
  Trash2,
} from 'lucide-react'
import {
  listboxDefinition,
  reducePatternData,
  useListboxPattern,
  type PatternData,
  type PatternEvent,
} from '@interactive-os/aria/react'
import { getRect, rectToStyle, type RetouchSlide } from './retouchModel'

type SlideRailProps = {
  activeSlideId: string
  canDeleteSlide: boolean
  canMoveSlideDown: boolean
  canMoveSlideUp: boolean
  canvasView: 'slide' | 'grid'
  changedSlideIds: Set<string>
  onAddSlide: () => void
  onChangeCanvasView: (view: 'slide' | 'grid') => void
  onDeleteSlide: () => void
  onDuplicateSlide: () => void
  onMoveSlideDown: () => void
  onMoveSlideUp: () => void
  onSelectSlide: (slideId: string) => void
  slides: RetouchSlide[]
}

export function SlideRail({
  activeSlideId,
  canDeleteSlide,
  canMoveSlideDown,
  canMoveSlideUp,
  canvasView,
  changedSlideIds,
  onAddSlide,
  onChangeCanvasView,
  onDeleteSlide,
  onDuplicateSlide,
  onMoveSlideDown,
  onMoveSlideUp,
  onSelectSlide,
  slides,
}: SlideRailProps) {
  const slideRailData = useMemo<PatternData>(() => {
    const slideIds = slides.map((slide) => slide.id)

    return {
      items: Object.fromEntries(
        slides.map((slide) => [
          slide.id,
          {
            label: changedSlideIds.has(slide.id)
              ? `${slide.name}, modified`
              : slide.name,
            textValue: slide.name,
          },
        ]),
      ),
      relations: { rootKeys: slideIds },
      state: {
        activeKey: activeSlideId,
        selectedKeys: [activeSlideId],
      },
      refs: { label: 'Slides' },
    }
  }, [activeSlideId, changedSlideIds, slides])

  const slideRailListbox = useListboxPattern(
    slideRailData,
    (event: PatternEvent) => {
      if (event.type === 'select') {
        const selectedSlideId = event.keys[0]

        if (selectedSlideId) {
          onSelectSlide(selectedSlideId)
        }
        return
      }

      if (event.type === 'navigate') {
        const nextSlideId = reducePatternData(
          listboxDefinition,
          slideRailData,
          event,
        ).state?.activeKey

        if (nextSlideId) {
          onSelectSlide(nextSlideId)
        }
      }
    },
    {
      elementIdPrefix: 'slide-thumb-',
      focusStrategy: 'rovingTabIndex',
      orientation: 'vertical',
      selectionMode: 'single',
      typeaheadEnabled: true,
    },
  )

  return (
    <aside className="slide-rail">
      <div className="view-toggle" role="tablist" aria-label="Canvas view">
        <button
          aria-selected={canvasView === 'slide'}
          onClick={() => onChangeCanvasView('slide')}
          role="tab"
          type="button"
        >
          Slide
        </button>
        <button
          aria-selected={canvasView === 'grid'}
          onClick={() => onChangeCanvasView('grid')}
          role="tab"
          type="button"
        >
          Grid
        </button>
      </div>

      <div className="rail-actions" role="toolbar" aria-label="Slide actions">
        <button aria-label="Add slide" onClick={onAddSlide} title="Add slide" type="button">
          <Plus aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Duplicate slide"
          onClick={onDuplicateSlide}
          title="Duplicate slide"
          type="button"
        >
          <Copy aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Move slide up"
          disabled={!canMoveSlideUp}
          onClick={onMoveSlideUp}
          title="Move slide up"
          type="button"
        >
          <ArrowUp aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Move slide down"
          disabled={!canMoveSlideDown}
          onClick={onMoveSlideDown}
          title="Move slide down"
          type="button"
        >
          <ArrowDown aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          aria-label="Delete slide"
          disabled={!canDeleteSlide}
          onClick={onDeleteSlide}
          title="Delete slide"
          type="button"
        >
          <Trash2 aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      </div>

      <div {...slideRailListbox.rootProps} className="slide-list">
        {slideRailListbox.renderItems.map((item, index) => {
          const slide = slides.find((candidate) => candidate.id === item.key)

          if (!slide) {
            return null
          }

          const changed = changedSlideIds.has(slide.id)
          const optionProps =
            item.optionProps as ButtonHTMLAttributes<HTMLButtonElement>

          return (
            <button
              {...optionProps}
              aria-current={slide.id === activeSlideId ? 'page' : undefined}
              aria-label={changed ? `${slide.name}, modified` : slide.name}
              className="slide-thumb"
              data-changed={changed ? 'true' : 'false'}
              key={slide.id}
              type="button"
            >
              <span className="thumb-number">{index + 1}</span>
              <MiniSlide slide={slide} />
              <span className="thumb-name">{slide.name}</span>
              {changed ? <span aria-hidden="true" className="thumb-change" /> : null}
            </button>
          )
        })}
      </div>
    </aside>
  )
}

function MiniSlide({ slide }: { slide: RetouchSlide }) {
  return (
    <span className="mini-slide" style={{ '--accent': slide.accent } as CSSProperties}>
      {slide.blocks.map((block) => (
        <span
          className={`mini-block ${block.role}`}
          key={block.id}
          style={rectToStyle(getRect(block))}
        />
      ))}
    </span>
  )
}
