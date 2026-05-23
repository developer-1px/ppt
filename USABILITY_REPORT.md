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
  - Fix: 원본 slide block 하나에만 `contenteditable="plaintext-only"`를 붙인다.
  - Fix: 편집 중에는 React가 block 내부 text children을 다시 렌더링하지 않게 하고, 브라우저 live DOM의 `textContent`를 draft로 사용한다.

- P0: Export 완료 액션 부재
  - Fix: Export panel에 `Copy` 버튼과 `Copied` feedback을 추가했다.
  - Fix: live draft나 drag draft가 보이는 동안에는 `Copied`/`Downloaded` feedback을 숨겨, 완료 상태가 현재 화면보다 앞서 보이지 않게 했다.
  - Fix: Clipboard와 fallback copy가 모두 실패하면 `Copied`로 표시하지 않고 실패 상태로 남긴다.

- P0: 편집 중 바로 Export하면 마지막 글자 수정이 누락됨
  - Fix: Copy/Download HTML은 먼저 live `contenteditable` DOM을 동기 commit하고, 그 최신 deck으로 export HTML을 만든다.

- P0: 편집 중 화면 이동으로 마지막 글자 수정이 누락될 수 있음
  - Fix: slide 전환과 Arrange 전환도 먼저 live `contenteditable` DOM을 commit한 뒤 이동한다.

- P0: 연속 텍스트 블록 편집이 검증되지 않음
  - Fix: 한 텍스트 블록 편집 중 다른 텍스트 블록을 클릭하면 첫 draft를 commit하고 두 번째 block 편집이 열리는 흐름을 verifier에 고정했다.

- P0: 편집 진입 시 실제 preview 텍스트와 editable 텍스트 사이에 visual gap이 생길 수 있음
  - Fix: `contenteditable` focus와 caret placement가 stage scroll을 움직이지 못하게 고정하고, compact viewport에서도 box/text 좌표가 변하지 않는지 verifier에 고정했다.
  - Fix: title token의 과도하게 낮은 `line-height`를 제거해 실제 글자 range가 preview/edit box 밖으로 튀어나가지 않게 했다.

- P0: 편집 중 toolbar Undo가 live draft를 건너뛸 수 있음
  - Fix: toolbar Undo/Redo도 먼저 live `contenteditable` DOM을 commit한 뒤 document history를 실행한다.

- P0: 편집 중 keyboard Undo가 live draft를 되돌리지 못함
  - Fix: Text Mode editor 안에서 `Cmd/Ctrl+Z`를 누르면 현재 live draft를 원래 문구로 되돌린다.

- P0: Text Mode에서 실수로 드래그해도 배치가 안 바뀌는지 약하게 검증됨
  - Fix: Text Mode drag attempt가 rect/text를 바꾸지 않고 resize handle도 노출하지 않는지 verifier에 고정했다.

- P0: First Experiment의 본문 문장 줄이기 플로우가 약하게 검증됨
  - Fix: `role="body"` 문장을 짧게 바꾸고, 위치/폭 유지와 Undo/Redo를 verifier에 고정했다.

- P0: 편집 중 Reset 후 Undo하면 live draft가 복구되지 않음
  - Fix: Text Mode reset도 먼저 live `contenteditable` DOM을 commit한 뒤 undoable root reset을 실행한다.

- P0: Text Mode Reset 범위가 전체 deck이라 한 문구 reset치고 위험함
  - Fix: Text Mode에서 block이 선택되어 있으면 Reset은 선택된 텍스트와 text-owned height만 원본으로 되돌린다.
  - Fix: 선택 텍스트 reset은 다른 layout/text 변경을 유지하고, Undo/Redo로 복구된다.

- P0: Text Mode에서 전체 deck reset 접근 경로가 선택 상태에 가려짐
  - Fix: 빈 stage를 클릭하면 Text Mode 선택이 해제되고, Reset이 전체 deck reset으로 바뀐다.

- P0: 편집 중 별도 박스 outline이 실제 글자와 떨어져 보임
  - Fix: Text Mode 편집 상태에서는 별도 boxed chrome을 그리지 않고 실제 글자 DOM과 caret만 사용한다.
  - Fix: Text Mode hover/focus도 일반 텍스트 block outline을 그리지 않는다.
  - Fix: 빈 텍스트 block만 다시 찾을 수 있도록 최소 dashed outline을 유지한다.

- P1: 직접 텍스트 편집에서 줄바꿈이 숨겨짐
  - Fix: plain Enter는 줄바꿈으로 두고, blur 또는 Cmd/Ctrl+Enter로 commit한다. Escape cancel은 유지한다.

- P1: Autoheight가 grow 중심이고 bottom clipping 정책이 없음
  - Fix: 편집 중에는 CSS `height: auto`로 자연스럽게 grow/shrink한다. commit 순간에만 DOM height를 측정해 저장하고, slide 하단을 넘으면 y를 보정해 box가 slide 안에 남게 한다.
  - Fix: 편집 overlay에서 block의 flex vertical centering을 reset해, 타이핑 중 content top이 재중앙정렬로 튀지 않게 했다.

- P1: Layout click selection / resize race
  - Fix: Layout Mode에서 block click도 selection으로 처리한다. verifier는 selected block 확인 후 resize한다.

- P1: Arrange Mode 중앙 정렬 기준 부족
  - Fix: block을 slide 중앙선 근처로 드래그하면 center snap이 적용되고 중앙 guide가 표시된다.

- P1: Arrange Mode가 slide 기준만 snap해서 카드끼리 정렬하기 어려움
  - Fix: 드래그 중인 block의 edge/center가 다른 block의 edge/center 근처에 오면 sibling snap guide가 적용된다.

- P1: Arrange Mode 미세 조정 부족
  - Fix: 선택된 block을 화살표 키로 한 칸씩 nudge할 수 있게 했다. Text Mode와 편집 중에는 동작하지 않는다.

- P1: 수정된 slide 식별 부족
  - Fix: 변경된 slide thumbnail에 작은 modified marker와 aria label을 추가했다.

- P1: slide 전환 후 stage scroll 잔류
  - Fix: slide 선택 시 stage scroll을 `(0, 0)`으로 reset한다.

- P1: Export style drift
  - Fix: export CSS에 preview block의 display, alignment, colors, padding, chart background를 반영했다.
  - Fix: slide 디자인 토큰을 `slideTheme.css`로 분리해 preview와 export가 같은 CSS 원천을 사용하게 했다.

- P1: Export가 최종 HTML만 만들고 retouch patch 근거를 잃음
  - Fix: exported HTML에 `data-retouch-patch` JSON manifest를 함께 포함해 text/layout 수정 내역을 구조화했다.

## zod-crud / Text Editing Usage

- `zod-crud`
  - `useJSONDocument`로 deck state, selection, history를 관리한다.
  - text/layout 변경은 JSON Patch로 commit한다.
  - Undo/Redo는 zod-crud history를 사용한다.

- `contenteditable="plaintext-only"`
  - Text Mode에서 실제 글자 편집 surface로 사용한다.
  - 직접 편집 중에는 별도 preview/editor block을 같이 렌더링하지 않는다.
  - 편집 중에는 비어 있지 않은 글자에 별도 boxed chrome을 그리지 않는다.
  - 편집 중에는 React children을 비워 live DOM과 preview state가 경쟁하지 않게 한다.
  - 별도 editor document를 만들지 않고 DOM textContent를 commit/cancel의 draft로 사용한다.
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
- Text Mode direct edit, Enter line break, shortcut/blur commit, Escape cancel
- no double text rendering
- live editor text and committed preview text parity without boxed editor chrome
- compact viewport edit entry without preview/editor visual gap
- preview/editor text range stays inside its box
- toolbar Undo/Redo around live text edits
- keyboard Undo inside live text edit
- Text Mode drag attempt does not change layout
- body sentence shortening without layout movement
- consecutive text block edits
- live edit commit before slide/mode switch
- selected text reset keeps other slide changes
- empty stage click exposes deck reset
- reset undo restores live text drafts
- autoheight grow/shrink, undo/redo, bottom slide fit
- Layout Mode center/sibling snap, arrow nudge, drag, resize, reset, no text editor
- Text Mode selected reset, undo reset, redo reset
- Export text/layout reflection, no editor chrome, Copy feedback
- Export feedback clears during live visible drafts
- Export Copy does not claim success when clipboard fails
- Export commits live text before Copy/Download
- Export shared slide theme tokens
- Export structured retouch patch manifest
- modified slide thumbnail state
- mobile horizontal overflow and core controls
