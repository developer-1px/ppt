# PPT UI core

PPT의 UI는 편집 엔진과 제품 셸을 분리한다.

## Ownership

- `core/`: 버튼, 탭, editor chrome 상태와 의미 토큰을 소유한다. PPT 모델이나 canvas 구현을 import하지 않는다.
- `shell/`: `EditorShell`, `EditorToolbar`, 상단 툴바, 슬라이드 레일, 인스펙터 배치와 점진 공개 규칙을 소유한다.
- `command-surface/`: PPT 명령 레지스트리, 표면별 명령 파생, 명령 버튼 표현을 소유한다.
- `inspector/`: Inspector가 소비하는 PPT 모델과 사용자 액션 계약, 패널별 편집 UI와 스타일을 소유한다. 앱은 탭 셸을 조합하고 액션을 도메인 명령으로 번역한다.
- `selection-toolbar/`: 선택 툴바의 조합, 점진 공개, 메뉴 상태, 빠른 텍스트 서식과 스타일을 소유한다. 외부에는 `model + onAction` 인터페이스만 노출한다.
- `text-formatting/`: 선택 툴바와 인스펙터가 함께 쓰는 PPT 텍스트 서식 제어를 소유한다.
- `App.css`: 슬라이드 내용과 캔버스 편집 affordance처럼 PPT 장면에만 해당하는 표현을 소유한다.
- `pptCanvas*Adapter.ts`: canvas 동작을 PPT 모델에 연결한다. UI core에 제품 동작을 밀어 넣지 않는다.

## Extension rules

1. 공통 컨트롤은 `core/index.ts`의 작은 공개 인터페이스를 통해 사용한다.
2. 아이콘 버튼은 `IconButton`의 `label`을 반드시 제공한다.
3. 툴바, 보기 옵션, 인스펙터 공개 상태는 독립 boolean 대신 `useEditorChrome`에서 관리한다.
4. 크기, 간격, radius, shadow, z-index는 `tokens.css`의 의미 토큰을 사용한다.
5. canvas가 제공하는 선택, 변형, 키보드 동작은 재구현하지 않는다.
6. 선택 툴바는 `model + onAction` 인터페이스로 조합하고, 제어별 콜백이나 공개 상태를 `App`으로 끌어올리지 않는다.
7. 앱 소유 UI의 레지스트리, 렌더링, 스타일은 같은 책임 폴더에 둔다.
8. Inspector는 필드별 콜백 대신 `PPTInspectorModel + PPTInspectorAction` 계약으로 연결한다.
9. Inspector 패널은 이 계약에서 필요한 모델만 읽고 같은 액션 스트림을 내보내며, 필드 계산과 이벤트 변환을 `App`으로 누출하지 않는다.
10. 책임 폴더가 있는 UI 스타일은 해당 폴더에서 로드하며 `App.css`에 다시 모으지 않는다.

`pnpm verify:ui-core`가 이 경계의 정적 가드다.
