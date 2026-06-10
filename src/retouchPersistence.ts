import { useEffect, useMemo, useState } from 'react'
import { createDirtyState } from '@zod-crud/dirty-state'
import { createDocumentPersistence } from '@zod-crud/persist-web'
import type { DocumentPersistenceCodec } from '@zod-crud/persist-web'
import type { JSONDocument } from 'zod-crud'
import { z } from 'zod'
import type { RetouchDeck } from './retouchModel'
import { retouchDeckEquals } from './retouchDirtyState'

const STORAGE_KEY = 'ppt-retouch:v3:deck'
const STORAGE_VERSION = 1
const RetouchDraftEnvelopeSchema = z.object({
  deck: z.unknown(),
  savedAt: z.unknown().optional(),
  selection: z.unknown().optional(),
  version: z.literal(STORAGE_VERSION),
})

const retouchDraftCodec: DocumentPersistenceCodec = {
  encode(input) {
    return JSON.stringify({
      version: STORAGE_VERSION,
      deck: input.value,
      selection: input.selection,
      savedAt: input.savedAt,
    })
  },

  decode(text) {
    const value: unknown = JSON.parse(text)
    const envelope = RetouchDraftEnvelopeSchema.safeParse(value)

    if (envelope.success) {
      return {
        value: envelope.data.deck,
        selection: null,
        savedAt:
          typeof envelope.data.savedAt === 'string' ? envelope.data.savedAt : null,
      }
    }

    return {
      value,
      selection: null,
      savedAt: null,
    }
  },
}

export function useRetouchDraftPersistence(doc: JSONDocument<RetouchDeck>) {
  const persistence = useMemo(
    () =>
      createDocumentPersistence(doc, {
        codec: retouchDraftCodec,
        key: STORAGE_KEY,
      }),
    [doc],
  )
  const dirtyState = useMemo(
    () =>
      createDirtyState(doc, {
        equals: retouchDeckEquals,
      }),
    [doc],
  )
  const [hasDeckChanges, setHasDeckChanges] = useState(() =>
    dirtyState.isDirty(),
  )
  const [persistenceReady, setPersistenceReady] = useState(false)

  useEffect(() => {
    const unsubscribe = dirtyState.subscribe((snapshot) => {
      setHasDeckChanges(snapshot.dirty)
    })

    return () => {
      unsubscribe()
    }
  }, [dirtyState])

  useEffect(() => {
    let cancelled = false

    void persistence
      .restore({
        restoreSelection: true,
      })
      .then(() => {
        if (!cancelled) {
          setHasDeckChanges(dirtyState.isDirty())
          setPersistenceReady(true)
        }
      })

    return () => {
      cancelled = true
    }
  }, [dirtyState, persistence])

  useEffect(() => {
    if (!persistenceReady) {
      return
    }

    void (dirtyState.isDirty() ? persistence.save() : persistence.clear())
  }, [dirtyState, doc.value, persistence, persistenceReady])

  return {
    hasDeckChanges,
    persistenceReady,
  }
}
