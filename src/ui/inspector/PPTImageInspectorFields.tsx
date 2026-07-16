import { ImagePlus, Undo2 } from 'lucide-react'
import { useRef } from 'react'
import {
  getPPTImageCrop,
  getPPTImageCropDescriptor,
  getPPTImageFit,
  getPPTImageReplaceDescriptor,
  isPPTImageFit,
} from '../../pptImageAdapter'
import { getPPTImageFileFromList } from '../../pptImportExtension'
import { Button } from '../core'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTImageInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTImageInspectorFields({
  model,
  onAction,
}: PPTImageInspectorFieldsProps) {
  const image = model.selectedElement
  const imageReplaceInputRef = useRef<HTMLInputElement | null>(null)
  const {
    onImageCropChange,
    onImageCropReset,
    onImageFitChange,
    onImageReplaceFile,
  } = createPPTInspectorActionDispatcher(onAction)

  if (image?.kind !== 'image') {
    return null
  }

  const imageCropDescriptor = getPPTImageCropDescriptor(model.slide.id, image)
  const imageReplaceDescriptor = getPPTImageReplaceDescriptor(
    model.slide.id,
    image,
  )

  return (
    <>
      <label className="ppt-field">
        <span>Fit</span>
        <select
          data-ppt-image-crop-attribute={imageCropDescriptor.metadata.attribute}
          data-ppt-image-crop-attribute-value={imageCropDescriptor.metadata.attributeValue}
          data-ppt-image-crop-command={imageCropDescriptor.fields.fit.commandId}
          data-ppt-image-crop-control={imageCropDescriptor.fields.fit.control}
          data-ppt-image-crop-field="fit"
          data-ppt-image-crop-supported={imageCropDescriptor.isSupported ? 'true' : 'false'}
          data-ppt-image-crop-surface={imageCropDescriptor.surface}
          data-ppt-style-field="image-fit"
          value={imageCropDescriptor.fit ?? getPPTImageFit(image)}
          onChange={(event) => {
            if (isPPTImageFit(event.target.value)) {
              onImageFitChange(image.id, event.target.value)
            }
          }}
        >
          {imageCropDescriptor.fields.fit.options.map((option) => (
            <option key={option.id} value={option.id}>
              {option.label}
            </option>
          ))}
        </select>
      </label>
      <div className="ppt-geometry-grid">
        {(['x', 'y'] as const).map((field) => {
          const imageCropField = imageCropDescriptor.fields[field]

          return (
            <label className="ppt-field" key={field}>
              <span>Crop {field.toUpperCase()}</span>
              <input
                data-ppt-image-crop-attribute={imageCropDescriptor.metadata.attribute}
                data-ppt-image-crop-attribute-value={imageCropDescriptor.metadata.attributeValue}
                data-ppt-image-crop-command={imageCropField.commandId}
                data-ppt-image-crop-control={imageCropField.control}
                data-ppt-image-crop-field={field}
                data-ppt-image-crop-supported={imageCropDescriptor.isSupported ? 'true' : 'false'}
                data-ppt-image-crop-surface={imageCropDescriptor.surface}
                data-ppt-image-crop-unit={imageCropField.unit}
                data-ppt-style-field={`image-crop-${field}`}
                max={imageCropField.max ?? 100}
                min={imageCropField.min ?? 0}
                step={imageCropField.step ?? 1}
                type="number"
                value={Math.round(imageCropDescriptor.crop[field] ?? getPPTImageCrop(image)[field])}
                onChange={(event) =>
                  onImageCropChange(
                    image.id,
                    field,
                    Number(event.target.value),
                  )}
              />
            </label>
          )
        })}
      </div>
      <Button
        data-ppt-image-crop-command={imageCropDescriptor.fields.reset.commandId}
        data-ppt-image-crop-control={imageCropDescriptor.fields.reset.control}
        data-ppt-image-crop-field="reset"
        data-ppt-image-crop-reset
        data-ppt-image-crop-supported={imageCropDescriptor.isSupported ? 'true' : 'false'}
        data-ppt-image-crop-surface={imageCropDescriptor.surface}
        disabled={imageCropDescriptor.isSupported === false}
        onClick={() => onImageCropReset(image.id)}
      >
        <Undo2 size={15} /> Reset
      </Button>
      <input
        accept={imageReplaceDescriptor.field.accept ?? 'image/*'}
        className="ppt-file-input"
        data-ppt-image-replace-attribute={imageReplaceDescriptor.metadata.attribute}
        data-ppt-image-replace-attribute-value={imageReplaceDescriptor.metadata.attributeValue}
        data-ppt-image-replace-command={imageReplaceDescriptor.field.commandId}
        data-ppt-image-replace-control={imageReplaceDescriptor.field.control}
        data-ppt-image-replace-field={imageReplaceDescriptor.field.id}
        data-ppt-image-replace-input
        data-ppt-image-replace-source-name={imageReplaceDescriptor.sourceName}
        data-ppt-image-replace-supported={imageReplaceDescriptor.isSupported ? 'true' : 'false'}
        data-ppt-image-replace-surface={imageReplaceDescriptor.surface}
        ref={imageReplaceInputRef}
        tabIndex={-1}
        type="file"
        onChange={(event) => {
          const file = getPPTImageFileFromList(event.target.files)

          if (file) {
            void onImageReplaceFile(image.id, file)
          }

          event.target.value = ''
        }}
      />
      <Button
        data-ppt-image-replace-action
        data-ppt-image-replace-command={imageReplaceDescriptor.field.commandId}
        data-ppt-image-replace-control={imageReplaceDescriptor.field.control}
        data-ppt-image-replace-field={imageReplaceDescriptor.field.id}
        data-ppt-image-replace-supported={imageReplaceDescriptor.isSupported ? 'true' : 'false'}
        data-ppt-image-replace-surface={imageReplaceDescriptor.surface}
        disabled={imageReplaceDescriptor.isSupported === false}
        onClick={() => imageReplaceInputRef.current?.click()}
      >
        <ImagePlus size={15} /> Change
      </Button>
    </>
  )
}
