# Usability Report

Date: 2026-05-23

## Scope

NORTH_STAR.md 기준으로 PPT retouch MVP를 검증했다.

- Text Mode: 글자만 직접 수정한다.
- Layout Mode: 선택, 드래그, 리사이즈, 리셋만 한다.
- Export: 현재 text/layout 상태가 반영된 HTML을 만든다.
- Autoheight: content에 맞춰 box height가 늘고 줄어든다. scroll/clip이 기본값이 아니다.

## Perspectives

### First-time user

- 첫 화면이 설명 페이지가 아니라 retouch editor인지 확인했다.
- Text/Layout 모드가 바로 보이고 전환 가능한지 확인했다.
- Text Mode에서 원본 텍스트와 editor가 이중 렌더링되지 않는지 확인했다.

### Designer / reviewer

- 정렬/배치가 편집 모드 진입으로 바뀌지 않는지 확인했다.
- Export가 preview와 크게 어긋나지 않도록 block style export를 보강했다.
- 하단 텍스트 autoheight가 slide 밖으로 잘려 보이지 않도록 확인했다.

### Rushed PPT owner

- Enter로 commit, Escape로 cancel, Undo/Redo로 복구되는지 확인했다.
- Layout Mode에서 텍스트가 열리지 않고 drag/resize/reset만 되는지 확인했다.
- Export 후 Copy 버튼과 completion feedback이 있는지 확인했다.

## Findings And Fixes

- P0: Text edit 이중 렌더링
  - Fix: 편집 중인 block은 preview block을 숨기고 `NanoTextEditor`만 렌더링한다.

- P0: Export 완료 액션 부재
  - Fix: Export panel에 `Copy` 버튼과 `Copied` feedback을 추가했다.

- P1: Enter commit 불일치
  - Fix: `NanoTextEditor`에서 native capture keydown으로 Enter commit, Escape cancel을 처리한다.

- P1: Autoheight가 grow 중심이고 bottom clipping 정책이 없음
  - Fix: 편집 중에는 CSS `height: auto`로 자연스럽게 grow/shrink한다. commit 순간에만 DOM height를 측정해 저장하고, slide 하단을 넘으면 y를 보정해 box가 slide 안에 남게 한다.
  - Fix: 편집 overlay에서 block의 flex vertical centering을 reset해, 타이핑 중 content top이 재중앙정렬로 튀지 않게 했다.

- P1: Layout click selection / resize race
  - Fix: Layout Mode에서 block click도 selection으로 처리한다. verifier는 selected block 확인 후 resize한다.

- P1: slide 전환 후 stage scroll 잔류
  - Fix: slide 선택 시 stage scroll을 `(0, 0)`으로 reset한다.

- P1: Export style drift
  - Fix: export CSS에 preview block의 display, alignment, colors, padding, chart background를 반영했다.
  - Residual: preview CSS와 export CSS는 아직 중복이다. 장기적으로 style token/source를 공유해야 한다.

## zod-crud / nano-edit Usage

- `zod-crud`
  - `useJSONDocument`로 deck state, selection, history를 관리한다.
  - text/layout 변경은 JSON Patch로 commit한다.
  - Undo/Redo는 zod-crud history를 사용한다.

- `nano-edit`
  - Text Mode에서 실제 글자 편집 engine/view로 사용한다.
  - 직접 편집 중에는 원본 preview block을 같이 렌더링하지 않는다.
  - PPT 도형 스타일 수치 편집 UI로 확장하지 않았다.

## Product Decision

Text Mode는 일반 layout tool을 제공하지 않는다.

다만 autoheight는 content-owned geometry로 본다. 즉 사용자가 글자를 바꾸면 box height가 content에 맞춰 바뀔 수 있고, 하단 clipping을 피하기 위한 최소 y 보정은 허용한다. 자유 이동/리사이즈는 Layout Mode에만 남긴다.

편집 중에는 React state로 height를 계속 쓰지 않는다. CSS autoheight로 보여주고, commit 시점에만 measured height를 zod-crud patch로 저장한다.

## Regression

Run:

```sh
pnpm lint
pnpm build
pnpm verify:mvp
```

`pnpm verify:mvp`는 매번 새 Vite port와 headless Chrome을 띄워 stale dev server 영향을 피한다.

Covered checks:

- first screen editor
- Text Mode direct edit, Enter commit, Escape cancel
- no double text rendering
- autoheight grow/shrink, undo/redo, bottom slide fit
- Layout Mode drag, resize, reset, no text editor
- Export text/layout reflection, no editor chrome, Copy feedback
- mobile horizontal overflow and core controls
