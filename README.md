# PPT

Lightweight PPT subset editor for AI-generated decks and final retouching.

## PPTX Viewer Goal

- Open a real `.pptx` file through file input or drop.
- Use the embedded PPT model when present; otherwise import from PPTX OpenXML.
- Render every imported slide as both a selectable thumbnail and the active slide page.
- Use `pnpm verify:pptx-render` as the fast gate for real-file page rendering.
- Without `PPTX_RENDER_FILE`, `pnpm verify:pptx-render` creates a temporary 3-slide `.pptx` file with slide backgrounds, theme-colored shapes, layout/master inherited marks, text, shapes, lines, hyperlinks, a grouped object, a table, and a cropped/flipped image on disk and opens it.
- Use `PPTX_RENDER_FILE=/path/to/file.pptx pnpm verify:pptx-render` to verify a provided PPTX file.
- Keep `pnpm verify:mvp` checking imported PPTX pages by iterating the rendered slides one by one.

## Architecture

- Source of truth: `PPTDeck`, `PPTSlide`, `PPTElement`
- Canvas usage: local PPT adapters wrap `canvas/core`, `canvas/foundation`, `canvas/app`, `canvas/engine`, and `canvas/renderer`
- Slide editing usage: `pptSlideEditAffordanceAdapter.ts` wraps the optional `@interactive-os/slide-edit-affordance` package
- DOM editing usage: `pptDomEditAffordanceAdapter.ts` wraps the optional `@interactive-os/dom-edit-affordance` metadata
- Product model: PPT types only. `CanvasItem` and `CanvasApp` stay package-side concepts.

## Run

```sh
pnpm install
pnpm dev
```

## Check

```sh
pnpm lint
pnpm build
pnpm verify:pptx-render
pnpm verify:mvp
```
