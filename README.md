# PPT

Lightweight PPT subset editor for AI-generated decks and final retouching.

## Architecture

- Source of truth: `PPTDeck`, `PPTSlide`, `PPTElement`
- Canvas usage: optional headless tools from `canvas/core` and `canvas/foundation`
- Not used as product model: `CanvasItem`, `CanvasApp`

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
