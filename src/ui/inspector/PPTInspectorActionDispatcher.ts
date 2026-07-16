import type {
  SlideEditColorSwatchSelection,
  SlideEditLayerPaneHostCommandEffect,
  SlideEditObjectVisibilityHostCommandEffect,
} from '../../pptSlideEditAffordanceAdapter'
import type {
  PPTElementAnimation,
  PPTFill,
  PPTImageFit,
  PPTLineMarker,
  PPTLineRoute,
  PPTParagraph,
  PPTShapeKind,
  PPTSlideTransition,
  PPTStroke,
  PPTTextStyle,
} from '../../pptModel'
import type {
  PPTInspectorColorSwatchChannel,
  PPTInspectorElementAnimationField,
  PPTInspectorElementShadowField,
  PPTInspectorParagraphSpacingField,
  PPTInspectorProps,
  PPTInspectorSlideTransitionField,
  PPTInspectorTextInsetField,
} from './PPTInspectorContract'

export function createPPTInspectorActionDispatcher(
  onAction: PPTInspectorProps['onAction'],
) {
  return {
    onColorSwatchApply(
      elementId: string,
      channel: PPTInspectorColorSwatchChannel,
      color: string,
      swatch: SlideEditColorSwatchSelection<string>,
    ) {
      onAction({
        channel,
        color,
        elementId,
        swatch,
        type: 'apply-color-swatch',
      })
    },
    onCommentBodyChange(elementId: string, body: string) {
      onAction({ body, elementId, type: 'set-comment-body' })
    },
    onCommentReplyAdd(elementId: string, reply: string) {
      onAction({ elementId, reply, type: 'add-comment-reply' })
    },
    onCommentResolvedChange(elementId: string, resolved: boolean) {
      onAction({ elementId, resolved, type: 'set-comment-resolved' })
    },
    onCommitText(elementId: string, text: string) {
      onAction({ elementId, text, type: 'commit-text' })
    },
    onCopyHTML() {
      onAction({ type: 'copy-html' })
    },
    onDownloadHTML() {
      onAction({ type: 'download-html' })
    },
    onElementAltTextChange(elementId: string, altText: string) {
      onAction({ altText, elementId, type: 'set-element-alt-text' })
    },
    onElementAnimationChange(
      elementId: string,
      field: PPTInspectorElementAnimationField,
      value: PPTElementAnimation[PPTInspectorElementAnimationField],
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-animation',
        value,
      })
    },
    onElementGeometryChange(
      elementId: string,
      field: 'h' | 'w' | 'x' | 'y',
      value: number,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-geometry',
        value,
      })
    },
    onElementHyperlinkChange(elementId: string, url: string) {
      onAction({ elementId, type: 'set-element-hyperlink', url })
    },
    onElementNameChange(elementId: string, name: string) {
      onAction({ elementId, name, type: 'set-element-name' })
    },
    onElementOpacityChange(elementId: string, opacity: number) {
      onAction({ elementId, opacity, type: 'set-element-opacity' })
    },
    onElementRotationChange(elementId: string, rotation: number) {
      onAction({ elementId, rotation, type: 'set-element-rotation' })
    },
    onElementShadowChange(
      elementId: string,
      field: PPTInspectorElementShadowField,
      value: boolean | number | string,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-shadow',
        value,
      })
    },
    onElementStrokeChange(
      elementId: string,
      field: keyof PPTStroke,
      value: number | string,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-stroke',
        value,
      })
    },
    onElementTextInsetChange(
      elementId: string,
      field: PPTInspectorTextInsetField,
      value: number,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-text-inset',
        value,
      })
    },
    onElementTextStyleChange(
      elementId: string,
      field: keyof PPTTextStyle,
      value: number | string,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-element-text-style',
        value,
      })
    },
    onImageCropChange(
      elementId: string,
      field: 'x' | 'y',
      value: number,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-image-crop',
        value,
      })
    },
    onImageCropReset(elementId: string) {
      onAction({ elementId, type: 'reset-image-crop' })
    },
    onImageFitChange(elementId: string, fit: PPTImageFit) {
      onAction({ elementId, fit, type: 'set-image-fit' })
    },
    onImageReplaceFile(
      elementId: string,
      file: Blob & { name?: string },
    ) {
      onAction({ elementId, file, type: 'replace-image-file' })
    },
    onLayerPaneCommandEffect(
      effect: SlideEditLayerPaneHostCommandEffect<string, string>,
    ) {
      onAction({ effect, type: 'run-layer-pane-command' })
    },
    onLayoutPlaceholderVisibilityChange(
      placeholderId: string,
      isVisible: boolean,
    ) {
      onAction({
        isVisible,
        placeholderId,
        type: 'set-placeholder-visibility',
      })
    },
    onLineMarkerChange(
      elementId: string,
      field: 'endMarker' | 'startMarker',
      marker: PPTLineMarker,
    ) {
      onAction({
        elementId,
        field,
        marker,
        type: 'set-line-marker',
      })
    },
    onLineRouteChange(elementId: string, route: PPTLineRoute) {
      onAction({ elementId, route, type: 'set-line-route' })
    },
    onObjectVisibilityCommandEffect(
      effect: SlideEditObjectVisibilityHostCommandEffect<string, string>,
    ) {
      onAction({ effect, type: 'run-object-visibility-command' })
    },
    onParagraphAlignChange(
      elementId: string,
      align: NonNullable<PPTParagraph['align']>,
    ) {
      onAction({ align, elementId, type: 'set-paragraph-align' })
    },
    onParagraphBulletChange(elementId: string, enabled: boolean) {
      onAction({ elementId, enabled, type: 'set-paragraph-bullet' })
    },
    onParagraphListLevelStep(elementId: string, delta: number) {
      onAction({ delta, elementId, type: 'step-paragraph-list-level' })
    },
    onParagraphNumberedChange(elementId: string, enabled: boolean) {
      onAction({ elementId, enabled, type: 'set-paragraph-numbered' })
    },
    onParagraphSpacingChange(
      elementId: string,
      field: PPTInspectorParagraphSpacingField,
      value: number,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-paragraph-spacing',
        value,
      })
    },
    onShapeCornerRadiusChange(
      elementId: string,
      cornerRadius: number,
    ) {
      onAction({
        cornerRadius,
        elementId,
        type: 'set-shape-corner-radius',
      })
    },
    onShapeFillChange(
      elementId: string,
      field: keyof PPTFill,
      value: number | string,
    ) {
      onAction({
        elementId,
        field,
        type: 'set-shape-fill',
        value,
      })
    },
    onShapeKindChange(elementId: string, shape: PPTShapeKind) {
      onAction({ elementId, shape, type: 'set-shape-kind' })
    },
    onSlideBackgroundChange(color: string) {
      onAction({ color, type: 'set-slide-background' })
    },
    onSlideLayoutChange(layoutId: string) {
      onAction({ layoutId, type: 'set-slide-layout' })
    },
    onSlideNameChange(name: string) {
      onAction({ name, type: 'set-slide-name' })
    },
    onSlideNotesChange(notes: string) {
      onAction({ notes, type: 'set-slide-notes' })
    },
    onSlideTransitionChange(
      field: PPTInspectorSlideTransitionField,
      value: PPTSlideTransition[PPTInspectorSlideTransitionField],
    ) {
      onAction({
        field,
        type: 'set-slide-transition',
        value,
      })
    },
    onTableRowsChange(elementId: string, rows: string) {
      onAction({ elementId, rows, type: 'set-table-rows' })
    },
    onTextAutoFit(elementId: string) {
      onAction({ elementId, type: 'auto-fit-text' })
    },
  }
}
