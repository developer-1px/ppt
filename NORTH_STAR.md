# North Star

## Product

AI가 만든 HTML/CSS PPT를 마지막에 아주 쉽게 리터칭하는 도구.

LLM 연동은 이미 끝났다고 가정한다. 이 앱은 생성기가 아니라 마감 편집기다. 사용자는 AI가 만든 슬라이드를 보고 글자를 조금 고치고, 배치를 조금 다듬고, 다시 HTML/CSS로 내보낸다.

## Core Bet

글자 수정과 배치 변경을 완전히 분리하면 AI 생성 PPT의 마지막 수정이 기존 PPT보다 쉽고 덜 불안해진다.

성공해야 하는 감각:

- 지금 가능한 조작이 즉시 이해된다.
- 글자를 고칠 때 배치가 망가질까 봐 불안하지 않다.
- 배치를 고칠 때 글자가 바뀔까 봐 불안하지 않다.
- CSS, 좌표, 폰트 수치 없이도 마지막 리터칭이 충분하다고 느낀다.

## MVP Goal

샘플 AI 생성 HTML/CSS 슬라이드를 로드하고, 사용자가 두 모드로만 마지막 수정을 끝낼 수 있게 한다.

- Text Mode: 문구만 수정한다.
- Layout Mode: 선택, 드래그, 리사이즈만 한다.
- Export: 원본 HTML/CSS와 patch를 합쳐 최종 HTML/CSS를 만든다.

## Non Goals

- LLM 연동을 만들지 않는다.
- CSS 에디터를 만들지 않는다.
- font-size, line-height, 좌표, width, height 같은 수치 입력을 노출하지 않는다.
- 색상 picker를 만들지 않는다.
- 복잡한 레이어 패널을 만들지 않는다.
- 템플릿/레이아웃 생성 UI를 만들지 않는다.
- Figma식 고급 도구를 만들지 않는다.
- 원본 HTML/CSS를 직접 계속 망가뜨리지 않는다.

## Mode Rules

### Text Mode

목적: 문구만 바꾼다.

허용:

- 텍스트 블록 클릭
- 직접 문구 수정
- Enter/Escape/blur로 확정 또는 취소
- text patch 저장

금지:

- 이동
- 리사이즈
- 스타일 수치 조정
- 배치 조정
- 디자인 패널 노출

### Layout Mode

목적: 글자를 건드리지 않고 배치만 다듬는다.

허용:

- 블록 클릭 선택
- 자유 드래그 이동
- 핸들 리사이즈
- snap guide
- 선택 블록 reset
- layout patch 저장

금지:

- 텍스트 직접 수정
- CSS 직접 수정
- 좌표 입력
- 크기 수치 입력
- 폰트 수치 입력

## Design System Rule

텍스트 스타일은 수치가 아니라 토큰으로만 선택한다.

허용되는 사고:

- 제목
- 부제목
- 본문
- 캡션
- 인용
- 강조
- 보조

금지되는 사고:

- 23px
- line-height 1.18
- weight 650
- letter-spacing
- arbitrary color

첫 MVP에서는 스타일 선택 UI도 필수가 아니다. 문구 수정과 배치 수정의 분리가 더 중요하다.

## Data Model

원본과 수정 patch를 분리한다.

```ts
type RetouchDeck = {
  slides: RetouchSlide[]
}

type RetouchSlide = {
  id: string
  html: string
  css: string
  textPatches: TextPatch[]
  layoutPatches: LayoutPatch[]
}

type TextPatch = {
  blockId: string
  text: string
}

type LayoutPatch =
  | { kind: 'translate'; blockId: string; x: number; y: number }
  | { kind: 'resize'; blockId: string; width: number; height: number }
  | { kind: 'reset'; blockId: string }
```

수치 입력은 UI에 노출하지 않는다. 내부 patch에는 렌더링과 export를 위해 수치를 저장할 수 있다.

## DOM Contract

AI 생성 HTML/CSS는 편집 가능한 최소 계약을 가져야 한다.

```html
<section data-slide="slide-1">
  <h1 data-block="title" data-role="title">Quarterly Update</h1>
  <p data-block="summary" data-role="body">...</p>
  <div data-block="card-1" data-role="card">...</div>
</section>
```

필수:

- `data-slide`
- `data-block`
- 안정적인 block id

권장:

- `data-role`
- 디자인 시스템 class

## UI Shape

첫 화면은 편집기여야 한다. 랜딩이나 설명 화면을 만들지 않는다.

```text
App
|
+-- Left: slide thumbnails
+-- Top: Text / Layout mode toggle
+-- Center: rendered HTML/CSS slide
+-- Overlay in Text Mode: text editing affordance only
+-- Overlay in Layout Mode: selection box, drag, resize handles, guides
+-- Bottom or compact side: undo, redo, reset, export
```

데모 설명과 장식은 극도로 절제한다.

## References To Follow

### ../nano-edit

본받을 것:

- 문구 수정 중심성
- patch와 undo/redo 사고방식
- 명령을 작고 명확하게 나누는 방식

그대로 가져오지 말 것:

- ProseMirror 중심의 복잡한 문서 편집기 표면
- 블록 편집기 전체 기능

### ../canvas

본받을 것:

- 선택 표시
- 자유 드래그
- 리사이즈 핸들
- snap guide
- renderer와 interaction 분리

그대로 가져오지 말 것:

- 전체 캔버스 제작 도구 표면
- 팔레트, 그룹, 정렬, 분배 같은 고급 도구를 MVP에 과하게 넣는 것

## First Experiment

가장 먼저 검증할 것:

```text
Text Mode와 Layout Mode를 분리하면
사용자가 AI 생성 PPT의 마지막 리터칭을 더 쉽게 끝낼 수 있는가?
```

검증 작업:

- 제목 문구 하나 바꾸기
- 본문 한 문장 줄이기
- 카드 하나 살짝 이동하기
- 카드 하나 살짝 키우기
- 실수 후 undo 또는 reset 하기
- 최종 HTML/CSS export 하기

## Auto Progression Rules

Goal mode로 자동 진행할 때는 아래 순서를 따른다.

1. Vite starter 화면을 제거하고 retouch shell을 만든다.
2. 샘플 AI 생성 HTML/CSS 슬라이드 2-3장을 코드에 내장한다.
3. slide thumbnail 전환을 만든다.
4. 중앙에 HTML/CSS slide preview를 렌더링한다.
5. Text Mode를 만든다.
6. text patch 저장과 undo/redo를 만든다.
7. Layout Mode를 만든다.
8. block selection, drag, resize handle을 만든다.
9. layout patch 저장과 selected block reset을 만든다.
10. export HTML/CSS를 만든다.
11. lint/build를 통과시킨다.
12. 브라우저에서 desktop/mobile 화면을 확인한다.

질문하지 않고 진행해도 되는 기본 판단:

- 샘플 데이터는 앱 내부 상수로 시작한다.
- LLM 입력창은 만들지 않는다.
- CSS 붙여넣기 입력창은 만들지 않는다.
- 배치 변경은 자유 드래그를 지원한다.
- 수치 입력은 만들지 않는다.
- 텍스트 스타일 UI는 첫 MVP에서 생략해도 된다.
- UI 문구는 최소화한다.

질문이 필요한 경우:

- LLM 연동을 실제로 붙이라는 요구가 생긴 경우
- 외부 HTML/CSS import를 MVP에 포함하라는 요구가 생긴 경우
- raw CSS 편집 기능을 넣자는 요구가 생긴 경우

## Definition Of Done

MVP가 끝났다고 볼 수 있는 상태:

- 앱 첫 화면이 retouch editor다.
- 사용자는 슬라이드를 전환할 수 있다.
- Text Mode에서 글자를 수정할 수 있다.
- Layout Mode에서 텍스트는 수정되지 않는다.
- Layout Mode에서 블록을 선택, 드래그, 리사이즈할 수 있다.
- 선택 블록의 layout patch를 reset할 수 있다.
- undo/redo가 동작한다.
- HTML/CSS export 결과가 현재 수정 상태를 반영한다.
- `pnpm lint`가 통과한다.
- `pnpm build`가 통과한다.
