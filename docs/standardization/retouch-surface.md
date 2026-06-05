# Retouch Surface Direction

## Purpose

PPT should stay a minimal dogfood host for final retouching of AI-generated
HTML/CSS slides. The repo should prove reusable parts, not grow into a full
presentation app or a full canvas editor.

Canonical seed: `src/retouchSurfaceContract.ts`.

## Evidence From Adjacent Packages

| Package | Stable responsibility | PPT implication |
|---|---|---|
| `zod-crud` | schema-backed JSON document, JSON Pointer/Patch, selection, history, extension labs | PPT document changes should be patch plans and commits, not ad-hoc React state mutations. |
| `../canvas` | foundation contracts for scene, selection, transform, snap, renderer adapters | PPT should reuse headless geometry and interaction contracts, not import the full CanvasApp surface. |
| `../apg-patterns` | serializable APG pattern runtime and React semantic props | PPT chrome should get roles, keyboard behavior, and state projection from APG patterns. Styling stays local. |
| `../spredsheet` | dogfood app that extracts contract seeds into `packages/*` and `docs/standardization` | PPT should record contract seeds before promoting packages. The host remains product-specific. |

## PPT-Owned Boundaries

PPT owns these because they define the product, not a reusable primitive:

- AI-generated HTML/CSS slide DOM contract: `data-slide`, `data-block`,
  `data-role`, stable block ids.
- Text Mode and Arrange Mode rules.
- Slide size, preview theme, and export shape for the current demo.
- Retouch patch manifest semantics: text patches and layout patches against a
  generated slide baseline.
- Minimal demo shell and verifier scenarios.

## Shared Boundaries To Prefer

Use existing packages when their semantics match:

- `zod-crud`: Retouch deck document, history, JSON Patch, Pointer, selection.
- `@zod-crud/collection`: slide array commands and block array commands.
- `@zod-crud/layer-order`: block z-order planning.
- `@zod-crud/dirty-state`: clean baseline comparison instead of local
  `JSON.stringify` equality helpers.
- `@zod-crud/persist-web`: browser draft save/restore instead of local storage
  envelope logic.
- `@zod-crud/id-resolver`: stable slide/block id to current Pointer lookup after
  reorder, duplicate, delete, or paste.
- `@interactive-os/aria`: tabs, toolbar, slide rail listbox/tree, dialogs, and
  future inspector controls.
- `@interactive-os/object-surface`: object selection, marquee, selection bounds.
- `canvas/core` and `canvas/foundation`: points, bounds, resize handles,
  drag threshold, scene adapter, transform, snap.

Do not force a package when the semantics do not match. Put an Adapter seam in
PPT first and record the mismatch.

## Candidate Interfaces

### Retouch Surface

The Retouch Surface candidate is the smallest portable contract for generated
slide retouching.

```ts
type RetouchSurface = {
  contract: 'interactive-os.retouch-surface.v1'
  schema: unknown
  slidePath: string
  slideSize: { width: number; height: number }
  blocksPath: string
  blockIdentity: { path: '/id' }
  capabilities: readonly (
    | 'text-edit'
    | 'autoheight'
    | 'object-selection'
    | 'marquee'
    | 'move'
    | 'resize'
    | 'z-order'
    | 'export-patch-manifest'
  )[]
}
```

The current PPT descriptor is `PPT_RETOUCH_SURFACE`. It fixes the v1 contract
name, deck schema, slide size, JSON Pointer templates, block identity path,
capabilities, and grouped JSON Patch runtime change shape.

Rules:

- Zod owns data shape and validation.
- The surface descriptor owns editable slide intent.
- Runtime output is grouped JSON Patch plus optional selection snapshot.
- Persistence, network, undo history, product chrome, and export UI stay outside
  the surface runtime.

### HTML Slide Contract

HTML Slide Contract should describe the minimum generated markup that a retouch
surface can index:

Canonical seed: `src/htmlSlideContract.ts`.

- one slide root per `data-slide`
- stable editable objects via `data-block`
- role hints via `data-role`
- explicit z-order via `data-block-index`
- contract id via `data-html-slide-contract`
- linked surface contract via `data-retouch-surface-contract`
- generated CSS may be preserved, but direct CSS editing is not part of this
  contract
- block text is plain text for the first contract version

The contract should not include LLM prompts, theme editors, template builders,
or a full presentation file format.

### Block Text Editor Adapter

Current pressure lives in `SlideBlockElement` and `editableTextDom`: IME gating,
`beforeinput`, caret placement, paste, trailing line breaks, autoheight, and
stage scroll preservation.

Current adapter seed: `src/plainTextBlockEditor.ts`.

Direction:

- Keep PPT-owned autoheight policy: text changes may change height; free
  layout movement remains Arrange Mode only.
- Use `@interactive-os/keyboard` for key predicates.
- Evaluate `@interactive-os/anyeditable` for plain-text contenteditable DOM
  ownership.
- If `anyeditable` is too composer-shaped, extract only the product-neutral DOM
  text editor kernel; keep slide-unit autoheight as a PPT Adapter.

### Slide Object Interaction Adapter

Current pressure lives in `layoutInteraction`, `useRetouchLayoutInteraction`,
`useRetouchMarqueeSelection`, and `retouchObjectSurface`.

Direction:

- Keep slide-unit projection local: client pixels to `1280x720` slide units.
- Use `object-surface` for selection and bounds.
- Adapt PPT blocks into `canvas/foundation` `CanvasSceneAdapter` entries.
- Reuse canvas transform/snap contracts only where PPT's text autoheight and
  slide boundary rules fit.
- Do not import CanvasApp. PPT is not a whiteboard product.

## Work Order

1. Replace local persistence and dirty helpers with `@zod-crud/persist-web` and
   `@zod-crud/dirty-state`.
2. Replace slide and block array commands with `@zod-crud/collection` where id
   generation remains host-owned. Current adapter: `src/retouchCollection.ts`
   handles slide move/delete and block delete; insert/duplicate stay PPT-owned
   until id resolver policy lands.
3. Add a Retouch id resolver Adapter around `@zod-crud/id-resolver`.
   Current adapter: `src/retouchIdResolver.ts` resolves slide/block ids through
   shared id-resolver and keeps PPT-owned id generation/duplicate policy in one
   place.
4. Add a Canvas scene Adapter for PPT blocks, then compare current move/snap
   behavior against `canvas/foundation`. Current adapter:
   `src/retouchCanvasScene.ts` maps PPT blocks to Canvas Foundation scene and
   transform contracts while keeping slide projection and autoheight in PPT.
5. Split block text editing into a plain-text editor Adapter and a PPT
   autoheight policy.
6. Expand APG usage from tabs to toolbar and slide rail selection.
7. Promote only after the gate in `package-promotion-gate.md` is satisfied by
   two hosts or one host plus a focused lab proving the same contract.

## Non-Goals

- Do not turn PPT into CanvasApp.
- Do not add spreadsheet grid semantics to PPT.
- Do not expose CSS, font size, line height, or arbitrary style editing.
- Do not create a package only to move one shallow helper.
- Do not hide product decisions behind generic names.
