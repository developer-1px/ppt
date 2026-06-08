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
  useActionToolbarPattern,
  useManagedListboxPattern,
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

const SLIDE_RAIL_ACTIONS = [
  { action: 'add', icon: Plus, label: 'Add slide' },
  { action: 'duplicate', icon: Copy, label: 'Duplicate slide' },
  { action: 'move-up', icon: ArrowUp, label: 'Move slide up' },
  { action: 'move-down', icon: ArrowDown, label: 'Move slide down' },
  { action: 'delete', icon: Trash2, label: 'Delete slide' },
] as const

type SlideRailActionKey = (typeof SLIDE_RAIL_ACTIONS)[number]['action']

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
  const railActionCommands = {
    add: onAddSlide,
    delete: onDeleteSlide,
    duplicate: onDuplicateSlide,
    'move-down': onMoveSlideDown,
    'move-up': onMoveSlideUp,
  } satisfies Record<SlideRailActionKey, () => void>
  const railActionDisabled = {
    add: false,
    delete: !canDeleteSlide,
    duplicate: false,
    'move-down': !canMoveSlideDown,
    'move-up': !canMoveSlideUp,
  } satisfies Record<SlideRailActionKey, boolean>
  const railActionToolbar = useActionToolbarPattern<SlideRailActionKey>({
    actions: SLIDE_RAIL_ACTIONS,
    disabledKeys: disabledToolbarKeys<SlideRailActionKey>(
      SLIDE_RAIL_ACTIONS.map(({ action }) => [
        action,
        railActionDisabled[action],
      ]),
    ),
    elementIdPrefix: 'slide-action-',
    label: 'Slide actions',
    onSelect: (action) => railActionCommands[action](),
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
        {SLIDE_RAIL_ACTIONS.map(({ action, icon: Icon, label }) => (
          <button
            {...railActionProps[action]}
            aria-label={label}
            disabled={railActionDisabled[action]}
            key={action}
            title={label}
            type="button"
          >
            <Icon aria-hidden="true" size={15} strokeWidth={2.2} />
          </button>
        ))}
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
