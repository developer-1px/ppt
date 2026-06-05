# Standardization Registry

This directory tracks package candidates discovered while building PPT retouch.
The app is a dogfood host first; a candidate belongs here only when the same
responsibility should survive outside this repo.

Promotion gate: `package-promotion-gate.md`.

| Candidate | Status | Source | Target package |
|---|---|---|---|
| Retouch surface | Contract seed | `src/retouchSurfaceContract.ts`, `src/retouchModel.ts`, `src/StageCanvas.tsx`, `src/retouchExport.ts` | `@interactive-os/retouch-surface` |
| HTML slide contract | Contract seed | `src/htmlSlideContract.ts`, `src/sampleDeck.ts`, `src/slideTheme.css`, `src/retouchExport.ts` | `@interactive-os/html-slide-contract` |
| Block text editor adapter | Pressure point | `src/SlideBlockElement.tsx`, `src/editableTextDom.ts` | `@interactive-os/anyeditable` adapter or `@interactive-os/block-text-editor` |
| Slide object interaction adapter | Pressure point | `src/layoutInteraction.ts`, `src/useRetouchLayoutInteraction.ts`, `src/retouchObjectSurface.ts` | `canvas/foundation` adapter |
| ARIA/APG usage | Shared package consumer | `src/Topbar.tsx`, `src/SlideRail.tsx`, `src/App.tsx` | `@interactive-os/aria` |
