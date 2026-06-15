# Standardization Registry

This repo now tracks the PPT subset boundary, not the old HTML retouch surface.

| Candidate | Status | Source | Target |
|---|---|---|---|
| PPT subset model | Seed | `src/pptModel.ts` | future PPTX bridge |
| PPT canvas adapter | Seed | `src/pptCanvasAdapter.ts` | `canvas/foundation` consumer pattern |
| PPT HTML export | Seed | `src/pptExport.ts` | temporary output path before PPTX export |

## Rules

- `PPTDeck` is the source of truth.
- `CanvasItem` is not the product model.
- `CanvasApp` is optional reference code, not the default app shell.
- Use `canvas/core` and `canvas/foundation` as headless tools only.
