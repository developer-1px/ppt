import { createCollection } from '@zod-crud/collection'
import type { CollectionEditResult } from '@zod-crud/collection'
import type { JSONDocument, Pointer } from 'zod-crud'
import {
  slidePointer,
  type RetouchDeck,
} from './retouchModel'

export function createRetouchCollection(doc: JSONDocument<RetouchDeck>) {
  const collection = createCollection(doc)

  return {
    deleteBlocks(pointers: readonly Pointer[]): CollectionEditResult {
      return collection.deleteItems(pointers)
    },

    deleteSlide(slideIndex: number): CollectionEditResult {
      return collection.deleteItems(slidePointer(slideIndex))
    },

    moveSlide(slideIndex: number, direction: -1 | 1): CollectionEditResult {
      return direction < 0
        ? collection.moveUp(slidePointer(slideIndex))
        : collection.moveDown(slidePointer(slideIndex))
    },

  }
}

export type RetouchCollection = ReturnType<typeof createRetouchCollection>
