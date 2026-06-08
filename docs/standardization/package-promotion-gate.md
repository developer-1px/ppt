# Retouch Package Candidate Promotion Gate

이 문서는 PPT 내부에서 생긴 후보를 공유 패키지로 승격할지 판단하는 기준이다.
목표는 PPT 앱을 키우는 것이 아니라, 재사용 가능한 부품을 검증해서 올바른
repo로 보내는 것이다.

## Gate 원칙

후보는 아래 조건을 만족할 때만 패키지 승격을 검토한다.

1. PPT 제품 정책과 공유 계약이 분리되어야 한다.
   예: Text Mode, Arrange Mode, sample deck, export UI는 PPT 정책이다. JSON
   Pointer, HTML data contract, APG props, scene adapter shape는 공유 계약
   후보가 될 수 있다.
2. 최소 두 사용처가 있거나, 한 사용처와 focused lab이 있어야 한다.
   PPT 하나만 쓰는 얕은 helper는 문서화된 seed로 둔다.
3. API는 versioned contract 또는 adapter surface로 고정되어야 한다.
   package entrypoint가 product 이름, demo 상태, React shell에 의존하면
   승격하지 않는다.
4. 테스트 seam이 있어야 한다.
   package 단위 테스트, PPT `verify:mvp`, 필요 시 target repo의 e2e/smoke가
   같은 실패를 잡아야 한다.
5. 승격 뒤 PPT가 package를 실제 소비해야 한다.
   복사본을 남기면 승격이 아니라 분기다.

## 단계

| 단계 | 의미 | 다음 단계 조건 |
|---|---|---|
| Seed | PPT 내부에 계약 이름, 타입, adapter가 생겼다. | 문서와 verifier가 같은 책임을 가리킨다. |
| Candidate | PPT 정책 밖으로 뺄 API 경계가 보인다. | 두 번째 host 또는 lab에서 같은 API가 필요하다. |
| Package-ready | target repo에서 테스트 가능한 surface가 있다. | target package export와 PPT 소비 PR을 함께 준비한다. |
| Promoted | 공유 repo가 source of truth가 되고 PPT는 import만 한다. | PPT 내부 중복 seed를 제거하거나 thin adapter로 축소한다. |

## 후보별 상태

### Retouch Surface

- 현재 상태: `src/retouchSurfaceContract.ts`에 v1 contract seed가 있다.
- target repo/package: 새 `@interactive-os/retouch-surface` 후보.
- PPT에 남길 것: slide size, demo deck, mode policy, export button, product
  naming.
- 공유 후보: surface id, slide/block pointer template, block identity,
  capability list, grouped JSON Patch runtime change shape.
- 승격 조건: PPT 외 host 또는 focused lab이 같은 descriptor를 소비하고,
  UI 없이 patch planning/identity lookup을 테스트할 수 있어야 한다.
- 다음 조치: retouch-surface lab을 만들거나 `../viewer`류 host에서 generated
  HTML retouch descriptor를 읽는 작은 소비자를 만든다.

### HTML Slide Contract

- 현재 상태: `src/htmlSlideContract.ts`가 canonical seed이고 preview/export DOM에
  `data-html-slide-contract`, `data-slide`, `data-block`, `data-role`,
  `data-block-index`를 붙인다.
- target repo/package: 새 `@interactive-os/html-slide-contract` 후보.
- PPT에 남길 것: 현재 slide theme, print CSS, patch manifest export shape.
- 공유 후보: DOM attribute names, root/block attr helpers. Snapshot reader와
  validator는 fixture 기반 검증이 생길 때 함께 추가한다.
- 승격 조건: fixture HTML을 package 단위에서 parse/validate하고, PPT export와
  다른 reader가 같은 index 결과를 써야 한다.
- 다음 조치: generated HTML fixture 기반 snapshot reader와 validator 테스트를
  추가하고, `../viewer` 또는 focused lab에서 읽기 전용 소비자를 검증한다.

### Block Text Editor Adapter

- 현재 상태: `src/plainTextBlockEditor.ts`가 contenteditable 속성, IME/beforeinput,
  paste, Enter, text normalization을 분리한다.
- target repo/package: 우선 `@interactive-os/anyeditable` 적합성 확인, 맞지
  않으면 `@interactive-os/block-text-editor` 후보.
- PPT에 남길 것: slide-unit autoheight, text block selection, commit/cancel
  policy, history commit timing.
- 공유 후보: plain text DOM editing kernel, Korean IME 중복 방지, paste
  normalization, caret insertion helper.
- 승격 조건: slide size와 React app 상태 없이 HTMLElement 단위 테스트로
  composition/input/paste/Enter가 검증되어야 한다.
- 다음 조치: Korean `insertText` regression과 composition path를 package test로
  옮길 수 있는 최소 DOM harness를 만든다.

### ARIA/APG Pattern Usage

- 현재 상태: sidebar view tabs, topbar toolbar, slide rail listbox/action toolbar가
  `@interactive-os/aria` React adapter를 사용한다.
- target repo/package: 기존 `../apg-patterns` (`@interactive-os/aria`).
- PPT에 남길 것: 버튼 배치, command 실행 함수, copy/download feedback state.
- 공유 후보: command toolbar에서 selectedKeys와 repeated action의 관계,
  roving focus props, listbox navigation semantics.
- 승격 조건: PPT에서 발견한 APG 동작 차이가 package-level test 또는 demo
  contract로 재현되어야 한다.
- 다음 조치: repeated command button은 `selectedKeys`에 묶지 않는다는 예시를
  `../apg-patterns` toolbar 문서나 테스트로 반영할지 판단한다.

### Canvas Adapter

- 현재 상태: `src/retouchCanvasScene.ts`가 PPT block을 `canvas/foundation`
  scene entry와 selection bounds로 변환한다.
- target repo/package: 기존 `../canvas` foundation/core.
- PPT에 남길 것: 1280x720 slide projection, text autoheight, slide boundary,
  selected block policy.
- 공유 후보: scene entry mapping, transform adapter, snap/resize input contract.
- 승격 조건: CanvasApp을 가져오지 않고 foundation contract만으로 move/resize/snap
  behavior를 비교 검증할 수 있어야 한다.
- 다음 조치: 현재 `layoutInteraction` 결과와 canvas foundation transform 결과를
  같은 fixture에서 비교하는 harness를 만든다.

### Zod CRUD Adapters

- 현재 상태: `@zod-crud/persist-web`, `@zod-crud/dirty-state`,
  `@zod-crud/collection`, `@zod-crud/id-resolver`를 PPT가 소비한다.
- target repo/package: 기존 `../zod-crud`.
- PPT에 남길 것: slide/block id 생성 정책, duplicate naming, reset scope UX.
- 공유 후보: array command, persistence envelope codec, dirty baseline,
  stable id resolver usage.
- 승격 조건: PPT에서 필요한 동작이 zod-crud package의 일반 document contract와
  맞고, package test로 잠글 수 있어야 한다.
- 다음 조치: `@zod-crud/layer-order`는 PPT와 `../canvas` 양쪽 사용처가 쌓이면
  lab에서 package 승격 여부를 판단한다.

## 검증 명령

PPT에서 후보를 seed/candidate로 인정하려면 아래를 통과해야 한다.

```sh
pnpm lint
pnpm exec tsc -b
pnpm build
pnpm verify:mvp
```

target repo로 승격할 때는 해당 repo의 검증도 같이 통과해야 한다.

| Repo | 최소 검증 | release급 검증 |
|---|---|---|
| `../apg-patterns` | `npm test`, `npm run typecheck` | `npm run check` |
| `../canvas` | `pnpm lint`, `pnpm test`, `pnpm build` | `pnpm test:e2e` 포함 |
| `../zod-crud` | `npm run extensions:verify` 또는 관련 workspace test | `npm run verify` |
| `../spredsheet` | package 추출 선례 비교 | `pnpm check` |

## 승격 차단 조건

- product demo 하나에서만 쓰는 얕은 helper다.
- sample deck, mode UX, CSS theme, export button 같은 PPT 정책이 섞여 있다.
- versioned contract 이름이 없다.
- package 단위에서 실패를 재현할 테스트 seam이 없다.
- target repo가 이미 같은 책임을 더 넓은 개념으로 소유한다.
- 승격 뒤에도 PPT 내부 구현이 source of truth로 남는다.
