# PPT UI core

PPT의 UI는 편집 엔진과 제품 셸을 분리한다.

## Ownership

- `core/`: 버튼, 탭, editor chrome 상태와 의미 토큰을 소유한다. PPT 모델이나 canvas 구현을 import하지 않는다.
- `shell/`: `EditorShell`, `EditorToolbar`, 상단 툴바, 슬라이드 레일, 인스펙터 배치와 점진 공개 규칙을 소유한다.
- `command-surface/`: PPT 명령 레지스트리, 표면별 명령 파생, 명령 버튼 표현을 소유한다.
- `inspector/`: Inspector가 소비하는 PPT 모델과 사용자 액션 계약, 탭·점진 공개·패널 조합, 패널별 편집 UI와 스타일을 소유한다. Objects 패널의 포커스, 이름 변경, 드래그, 키보드 상태도 이 경계 안에 둔다. 앱은 액션을 도메인 명령으로 번역한다.
- `selection-toolbar/`: 선택 툴바의 조합, 점진 공개, 메뉴 상태, 빠른 텍스트 서식과 스타일을 소유한다. 외부에는 `model + onAction` 인터페이스만 노출한다.
- `text-formatting/`: 선택 툴바와 인스펙터가 함께 쓰는 PPT 텍스트 서식 제어를 소유한다.
- `App.css`: 슬라이드 내용과 캔버스 편집 affordance처럼 PPT 장면에만 해당하는 표현을 소유한다.
- `pptCanvas*Adapter.ts`: canvas 동작을 PPT 모델에 연결한다. UI core에 제품 동작을 밀어 넣지 않는다.
- `pptLayerPaneAdapter.ts`: PPT 요소와 그룹을 canvas Layer Pane descriptor, 선택, 재정렬 명령으로 변환한다. 앱의 도메인 명령과 Inspector UI가 함께 쓰는 PPT 경계다.
- `pptCommentThreadAdapter.ts`: comment 본문과 reply를 정규화하고 저장된 thread가 없을 때 표시할 fallback thread를 만든다.
- `pptImageAdapter.ts`: 이미지 fit/crop 기본값을 정규화하고 canvas Image Crop/Replace descriptor로 변환한다. 캔버스 렌더링과 Inspector가 같은 PPT 이미지 값을 사용하게 하는 경계다.
- `pptObjectAdapter.ts`: 모든 PPT 객체에 공통인 opacity, hyperlink, accessibility, shadow 값을 정규화하고 canvas descriptor로 변환한다. 렌더링·명령 처리·Inspector가 같은 객체 의미를 사용하게 하는 경계다.
- `pptObjectAnimationAdapter.ts`: PPT 객체 애니메이션의 기본값, 정규화, JSON 판별, canvas 명령·descriptor·build order·CSS projection을 소유한다.
- `pptColorSwatchAdapter.ts`: Inspector의 색상 target, PPT 채널과 canvas 채널 매핑, theme/recent palette descriptor를 소유한다.
- `pptTextAdapter.ts`: PPT 텍스트 스타일, 문단, font/vertical/inset descriptor, autofit size mode와 canvas 명령·CSS projection을 소유한다.

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
11. Objects 패널의 이름 변경, 드래그, 포커스 같은 일시적 상태는 `PPTLayerPane`이 소유하고 `App`에는 `model + onAction`만 노출한다.
12. Inspector 탭, 선택 변화에 따른 활성 탭, 고급 속성 공개 상태와 Slide/Objects/Export 조합은 `PPTInspectorShell`이 소유한다.
13. Comment reply draft와 commit interaction은 `PPTCommentInspectorFields`가 소유하며, comment 저장 모델 정규화는 `pptCommentThreadAdapter.ts`를 함께 사용한다.
14. Image Inspector의 crop/fit/replace 필드와 파일 input ref는 `PPTImageInspectorFields`가 소유하며, 이미지 값과 descriptor 변환은 `pptImageAdapter.ts`를 함께 사용한다.
15. 공통 Object Properties의 name, opacity, hyperlink, accessibility, shadow, geometry 필드는 `PPTObjectPropertiesInspectorFields`가 소유하며, 객체 값과 descriptor 변환은 `pptObjectAdapter.ts`를 함께 사용한다.
16. Object Animation 필드와 선택값 검증은 `PPTObjectAnimationInspectorFields`가 소유하며, 애니메이션 모델·명령·descriptor 변환은 `pptObjectAnimationAdapter.ts`를 함께 사용한다.
17. Color Swatch Strip은 callback props 대신 `model + onAction + target` 계약을 사용하고, 채널·palette 변환은 `pptColorSwatchAdapter.ts`를 함께 사용한다.
18. Text Inspector의 text/font/paragraph/inset/autofit 필드는 `PPTTextInspectorFields`가 소유하며, 텍스트 값·명령·descriptor 변환은 `pptTextAdapter.ts`를 함께 사용한다.

`pnpm verify:ui-core`가 이 경계의 정적 가드다.
