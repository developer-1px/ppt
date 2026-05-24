import { type ButtonHTMLAttributes, type CSSProperties, useMemo } from 'react'
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
  canvasView: 'slide' | 'grid'
  changedSlideIds: Set<string>
  onChangeCanvasView: (view: 'slide' | 'grid') => void
  onSelectSlide: (slideId: string) => void
  slides: RetouchSlide[]
}

export function SlideRail({
  activeSlideId,
  canvasView,
  changedSlideIds,
  onChangeCanvasView,
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
