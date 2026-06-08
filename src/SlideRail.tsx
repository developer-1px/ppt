import {
  type ButtonHTMLAttributes,
  type HTMLAttributes,
  useMemo,
} from 'react'
import {
  ArrowDown,
  ArrowUp,
  Copy,
  Plus,
  Trash2,
} from 'lucide-react'
import { getRect, rectToStyle, type RetouchSlide } from './retouchModel'
import {
  disabledToolbarKeys,
  useManagedListboxPattern,
  useManagedToolbarPattern,
} from './apgPatternAdapter'
import { cssVariables } from './cssVariables'
import './SlideRail.css'

type SlideRailProps = {
  activeSlideId: string
  canDeleteSlide: boolean
  canMoveSlideDown: boolean
  canMoveSlideUp: boolean
  canvasViewTablistProps: HTMLAttributes<HTMLDivElement>
  canvasViewTabProps: Record<'slide' | 'grid', ButtonHTMLAttributes<HTMLButtonElement>>
  changedSlideIds: Set<string>
  onAddSlide: () => void
  onDeleteSlide: () => void
  onDuplicateSlide: () => void
  onMoveSlideDown: () => void
  onMoveSlideUp: () => void
  onSelectSlide: (slideId: string) => void
  slides: RetouchSlide[]
}

type SlideRailActionKey =
  | 'add'
  | 'delete'
  | 'duplicate'
  | 'move-down'
  | 'move-up'

const SLIDE_RAIL_ACTION_KEYS = [
  'add',
  'duplicate',
  'move-up',
  'move-down',
  'delete',
] as const

export function SlideRail({
  activeSlideId,
  canDeleteSlide,
  canMoveSlideDown,
  canMoveSlideUp,
  canvasViewTablistProps,
  canvasViewTabProps,
  changedSlideIds,
  onAddSlide,
  onDeleteSlide,
  onDuplicateSlide,
  onMoveSlideDown,
  onMoveSlideUp,
  onSelectSlide,
  slides,
}: SlideRailProps) {
  const slideRailKeys = useMemo(() => slides.map((slide) => slide.id), [slides])
  const slideRailItems = useMemo(() => Object.fromEntries(
    slides.map((slide) => [
      slide.id,
      {
        label: changedSlideIds.has(slide.id)
          ? `${slide.name}, modified`
          : slide.name,
        textValue: slide.name,
      },
    ]),
  ), [changedSlideIds, slides])
  const slidesById = useMemo(
    () => new Map(slides.map((slide) => [slide.id, slide])),
    [slides],
  )
  const slideRailListbox = useManagedListboxPattern<string>({
    activeKey: activeSlideId,
    elementIdPrefix: 'slide-thumb-',
    items: slideRailItems,
    label: 'Slides',
    onSelect: onSelectSlide,
    rootKeys: slideRailKeys,
  })
  const railActionToolbar = useManagedToolbarPattern<SlideRailActionKey>({
    disabledKeys: disabledToolbarKeys<SlideRailActionKey>([
      ['move-up', !canMoveSlideUp],
      ['move-down', !canMoveSlideDown],
      ['delete', !canDeleteSlide],
    ]),
    elementIdPrefix: 'slide-action-',
    handlers: {
      add: onAddSlide,
      delete: onDeleteSlide,
      duplicate: onDuplicateSlide,
      'move-down': onMoveSlideDown,
      'move-up': onMoveSlideUp,
    },
    items: {
      add: { label: 'Add slide' },
      delete: { label: 'Delete slide' },
      duplicate: { label: 'Duplicate slide' },
      'move-down': { label: 'Move slide down' },
      'move-up': { label: 'Move slide up' },
    },
    label: 'Slide actions',
    omitPressed: true,
    rootKeys: SLIDE_RAIL_ACTION_KEYS,
  })
  const railActionProps = railActionToolbar.itemProps

  return (
    <aside className="slide-rail">
      <div {...canvasViewTablistProps} className="view-toggle">
        <button
          {...canvasViewTabProps.slide}
          type="button"
        >
          Slide
        </button>
        <button
          {...canvasViewTabProps.grid}
          type="button"
        >
          Grid
        </button>
      </div>

      <div {...railActionToolbar.rootProps} className="rail-actions">
        <button
          {...railActionProps.add}
          aria-label="Add slide"
          title="Add slide"
          type="button"
        >
          <Plus aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps.duplicate}
          aria-label="Duplicate slide"
          title="Duplicate slide"
          type="button"
        >
          <Copy aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps['move-up']}
          aria-label="Move slide up"
          disabled={!canMoveSlideUp}
          title="Move slide up"
          type="button"
        >
          <ArrowUp aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps['move-down']}
          aria-label="Move slide down"
          disabled={!canMoveSlideDown}
          title="Move slide down"
          type="button"
        >
          <ArrowDown aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
        <button
          {...railActionProps.delete}
          aria-label="Delete slide"
          disabled={!canDeleteSlide}
          title="Delete slide"
          type="button"
        >
          <Trash2 aria-hidden="true" size={15} strokeWidth={2.2} />
        </button>
      </div>

      <div {...slideRailListbox.rootProps} className="slide-list">
        {slideRailListbox.renderItems.map((item, index) => {
          const slide = slidesById.get(item.key)

          if (!slide) {
            return null
          }

          const changed = changedSlideIds.has(slide.id)

          return (
            <button
              {...item.optionProps}
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
    <span className="mini-slide" style={cssVariables({ '--accent': slide.accent })}>
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
