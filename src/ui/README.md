# PPT UI core

PPT의 UI는 편집 엔진과 제품 셸을 분리한다.

## Ownership

- `core/`: 버튼, 탭, editor chrome 상태와 의미 토큰을 소유한다. PPT 모델이나 canvas 구현을 import하지 않는다.
- `shell/`: `EditorShell`, `EditorToolbar`, 상단 툴바, 슬라이드 레일, 인스펙터 배치와 점진 공개 규칙을 소유한다.
- `selection-toolbar/`: 선택 툴바의 메뉴 레지스트리, 공개 상태, 키보드 포커스, 미리보기 수명을 소유한다. `App`에는 도메인 명령과 실행 가능한 상태만 요구한다.
- `App.css`: 슬라이드 내용, 편집 affordance, 도메인별 인스펙터 필드처럼 PPT에만 해당하는 표현을 소유한다.
- `pptCanvas*Adapter.ts`: canvas 동작을 PPT 모델에 연결한다. UI core에 제품 동작을 밀어 넣지 않는다.

## Extension rules

1. 공통 컨트롤은 `core/index.ts`의 작은 공개 인터페이스를 통해 사용한다.
2. 아이콘 버튼은 `IconButton`의 `label`을 반드시 제공한다.
3. 툴바, 보기 옵션, 인스펙터 공개 상태는 독립 boolean 대신 `useEditorChrome`에서 관리한다.
4. 크기, 간격, radius, shadow, z-index는 `tokens.css`의 의미 토큰을 사용한다.
5. canvas가 제공하는 선택, 변형, 키보드 동작은 재구현하지 않는다.
6. 선택 툴바 메뉴는 `selection-toolbar/index.ts`의 공개 인터페이스로 조합하고, 메뉴 내부 상태를 `App`으로 끌어올리지 않는다.

`pnpm verify:ui-core`가 이 경계의 정적 가드다.
