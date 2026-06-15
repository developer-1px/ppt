# North Star

## Product

AI가 생성한 최소 PPT subset을 사용자가 마지막에 가볍게 리터치하거나,
간단한 PPT를 직접 만들 수 있는 편집기.

PPTX export/import는 장기 방향이다. 지금은 PPTX로 갈 수 있는 자체 모델을
먼저 세우고, HTML export로 현재 결과를 확인한다.

## Core Bet

원본 데이터를 `PPT*` 모델로 소유하고, 선택/드래그/리사이즈/snap 같은
조작만 `canvas` 패키지 도구로 위임하면, 가벼운 PPT 제작과 AI 결과물
리터치를 동시에 만족할 수 있다.

## Model Rule

`PPT` 표기를 사용한다.

```ts
type PPTDeck = {
  id: string
  title: string
  size: { w: number; h: number }
  slides: PPTSlide[]
}

type PPTSlide = {
  id: string
  name: string
  background?: PPTFill
  elements: PPTElement[]
  notes?: string
}

type PPTElement = PPTTextBox | PPTShape | PPTImage
```

텍스트는 모델에서 `textBody.paragraphs.runs` 형태를 유지한다. UI는 당분간
plain text처럼 편집해도 된다.

## Canvas Rule

`canvas`는 패키지 도구다.

사용:

- `canvas/core`: bounds, viewport, resize handles, zoom primitives
- `canvas/foundation`: scene adapter, selection, marquee, transform, snap

사용하지 않음:

- `CanvasItem`을 PPT 원본 모델로 사용하지 않는다.
- `CanvasApp`을 앱의 기본 shell로 사용하지 않는다.

## MVP

- 슬라이드 rail
- 16:9 stage
- `PPTElement` 선택
- drag/resize/marquee
- text retouch
- undo/redo
- HTML export with embedded `PPTDeck` JSON

## Non Goals

- full PowerPoint 재현
- PPTX export/import 구현
- LLM 연동
- Figma식 고급 도구
- raw CSS 편집
- `CanvasItem` 중심 앱 모델

## Verification

```sh
pnpm lint
pnpm build
pnpm verify:mvp
```
