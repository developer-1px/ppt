import type {
  SlideEditInspectorSurfaceId,
  SlideEditLayerPaneHostCommandEffect,
  SlideEditLayoutDescriptor,
  SlideEditLayoutPlaceholderVisibilityHostCommandEffect,
  SlideEditObjectVisibilityHostCommandEffect,
  SlideEditPlaceholderDescriptor,
  SlideEditResolvedLayoutPlaceholder,
  SlideEditSlideMetadataInspectorDescriptor,
  SlideEditSlideMetadataReadModel,
  SlideEditSlideOrientation,
  SlideEditSlideSizeDescriptor,
  SlideEditTextAutoFitHostCommandEffect,
  SlideEditTextOverflowIndicatorState,
  SlideEditThemeColorToken,
} from '../../pptSlideEditAffordanceAdapter'
import type {
  PPTElement,
  PPTElementAnimation,
  PPTFill,
  PPTImageFit,
  PPTLineMarker,
  PPTLineRoute,
  PPTParagraph,
  PPTShapeKind,
  PPTSlide,
  PPTSlideTransition,
  PPTStroke,
  PPTTextStyle,
} from '../../pptModel'
import type { PPTElementShadowUpdateField } from '../../pptObjectAdapter'
import type { PPTElementAnimationUpdateField } from '../../pptObjectAnimationAdapter'
import type {
  PPTColorSwatchChannel,
  PPTColorSwatchSelection,
} from '../../pptColorSwatchAdapter'
import type {
  PPTParagraphSpacingField,
  PPTTextInsetField,
} from '../../pptTextAdapter'

export type PPTInspectorSurfaceId = SlideEditInspectorSurfaceId

type PPTInspectorSlideMetadataDescriptor =
  SlideEditSlideMetadataInspectorDescriptor<string> & {
    metadata: SlideEditSlideMetadataReadModel<string> & {
      orientation: SlideEditSlideOrientation
      size: SlideEditSlideSizeDescriptor
    }
  }

export type PPTInspectorColorSwatchChannel = PPTColorSwatchChannel

export type PPTInspectorElementAnimationField = PPTElementAnimationUpdateField

export type PPTInspectorElementShadowField = PPTElementShadowUpdateField

export type PPTInspectorParagraphSpacingField = PPTParagraphSpacingField

export type PPTInspectorSlideTransitionField =
  | 'advanceAfterMs'
  | 'advanceOnClick'
  | 'durationMs'
  | 'type'

export type PPTInspectorTextInsetField = PPTTextInsetField

export type PPTInspectorModel = {
  exportCode: string
  hidden: boolean
  inspectorSurface: PPTInspectorSurfaceId
  layoutDescriptors: readonly SlideEditLayoutDescriptor[]
  layoutPlaceholderVisibilityDescriptors: readonly SlideEditPlaceholderDescriptor<string, string>[]
  layoutPlaceholders: readonly SlideEditResolvedLayoutPlaceholder[]
  lastPlaceholderVisibilityEffect:
    SlideEditLayoutPlaceholderVisibilityHostCommandEffect<string, string> | null
  lastTextAutoFitEffect: SlideEditTextAutoFitHostCommandEffect<string, string> | null
  recentColors: readonly string[]
  selection: string[]
  selectedElement: PPTElement | null
  selectedElementAnimation: PPTElementAnimation | null
  selectedTextOverflow: boolean
  slide: PPTSlide
  slideLayoutId: string
  slideMetadataDescriptor: PPTInspectorSlideMetadataDescriptor
  slideThemeId: string
  slideTransition: PPTSlideTransition
  textAutoFitIndicator: SlideEditTextOverflowIndicatorState<string, string> | null
  themeColorTokens: readonly SlideEditThemeColorToken[]
}

export type PPTInspectorAction =
  | {
      channel: PPTInspectorColorSwatchChannel
      color: string
      elementId: string
      swatch: PPTColorSwatchSelection
      type: 'apply-color-swatch'
    }
  | { elementId: string; type: 'auto-fit-text' }
  | { elementId: string; text: string; type: 'commit-text' }
  | { type: 'copy-html' }
  | { type: 'download-html' }
  | { altText: string; elementId: string; type: 'set-element-alt-text' }
  | {
      elementId: string
      field: PPTInspectorElementAnimationField
      type: 'set-element-animation'
      value: PPTElementAnimation[PPTInspectorElementAnimationField]
    }
  | {
      elementId: string
      field: 'h' | 'w' | 'x' | 'y'
      type: 'set-element-geometry'
      value: number
    }
  | { elementId: string; type: 'set-element-hyperlink'; url: string }
  | { elementId: string; name: string; type: 'set-element-name' }
  | { elementId: string; opacity: number; type: 'set-element-opacity' }
  | { elementId: string; rotation: number; type: 'set-element-rotation' }
  | {
      elementId: string
      field: PPTInspectorElementShadowField
      type: 'set-element-shadow'
      value: boolean | number | string
    }
  | {
      elementId: string
      field: keyof PPTStroke
      type: 'set-element-stroke'
      value: number | string
    }
  | {
      elementId: string
      field: PPTInspectorTextInsetField
      type: 'set-element-text-inset'
      value: number
    }
  | {
      elementId: string
      field: keyof PPTTextStyle
      type: 'set-element-text-style'
      value: number | string
    }
  | { elementId: string; type: 'set-image-crop'; field: 'x' | 'y'; value: number }
  | { elementId: string; fit: PPTImageFit; type: 'set-image-fit' }
  | {
      elementId: string
      file: Blob & { name?: string }
      type: 'replace-image-file'
    }
  | { elementId: string; type: 'reset-image-crop' }
  | {
      effect: SlideEditLayerPaneHostCommandEffect<string, string>
      type: 'run-layer-pane-command'
    }
  | {
      effect: SlideEditObjectVisibilityHostCommandEffect<string, string>
      type: 'run-object-visibility-command'
    }
  | {
      elementId: string
      field: 'endMarker' | 'startMarker'
      marker: PPTLineMarker
      type: 'set-line-marker'
    }
  | { elementId: string; route: PPTLineRoute; type: 'set-line-route' }
  | { elementId: string; enabled: boolean; type: 'set-paragraph-bullet' }
  | { elementId: string; enabled: boolean; type: 'set-paragraph-numbered' }
  | {
      align: NonNullable<PPTParagraph['align']>
      elementId: string
      type: 'set-paragraph-align'
    }
  | {
      delta: number
      elementId: string
      type: 'step-paragraph-list-level'
    }
  | {
      elementId: string
      field: PPTInspectorParagraphSpacingField
      type: 'set-paragraph-spacing'
      value: number
    }
  | { body: string; elementId: string; type: 'set-comment-body' }
  | { elementId: string; resolved: boolean; type: 'set-comment-resolved' }
  | { elementId: string; reply: string; type: 'add-comment-reply' }
  | { cornerRadius: number; elementId: string; type: 'set-shape-corner-radius' }
  | {
      elementId: string
      field: keyof PPTFill
      type: 'set-shape-fill'
      value: number | string
    }
  | { elementId: string; shape: PPTShapeKind; type: 'set-shape-kind' }
  | { color: string; type: 'set-slide-background' }
  | { layoutId: string; type: 'set-slide-layout' }
  | { name: string; type: 'set-slide-name' }
  | { notes: string; type: 'set-slide-notes' }
  | {
      field: PPTInspectorSlideTransitionField
      type: 'set-slide-transition'
      value: PPTSlideTransition[PPTInspectorSlideTransitionField]
    }
  | { isVisible: boolean; placeholderId: string; type: 'set-placeholder-visibility' }
  | { elementId: string; rows: string; type: 'set-table-rows' }

export type PPTInspectorProps = {
  model: PPTInspectorModel
  onAction: (action: PPTInspectorAction) => void
}
