# Usability Report

Date: 2026-06-15

## Scope

기존 HTML/CSS retouch 앱을 제거하고, 자체 `PPT*` subset 모델 기반 편집기로
재구성했다.

검증 대상:

- 첫 화면이 PPT subset editor다.
- `CanvasItem`/`CanvasApp`이 제품 모델로 노출되지 않는다.
- `PPTElement` 선택과 drag가 동작한다.
- drag는 undo history에 기록된다.
- HTML export가 slide/element markup과 embedded `PPTDeck` JSON을 포함한다.
- mobile viewport에서 stage가 유지된다.

## Result

`pnpm verify:mvp`가 headless Chrome으로 위 시나리오를 검증한다.
