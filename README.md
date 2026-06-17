# PPT

Lightweight PPT subset editor for AI-generated decks and final retouching.

## Architecture

- Source of truth: `PPTDeck`, `PPTSlide`, `PPTElement`
- Canvas usage: local PPT adapters wrap `canvas/core`, `canvas/foundation`, `canvas/app`, `canvas/engine`, and `canvas/renderer`
- Slide editing usage: `pptSlideEditAffordanceAdapter.ts` wraps the optional `@interactive-os/slide-edit-affordance` package
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
pnpm verify:mvp
```
