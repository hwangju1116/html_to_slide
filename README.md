# HTML → PPTX 변환 스킬 (`html-to-pptx`)

HTML 슬라이드와 슬라이드 이미지를 **편집 가능한 16:9 PowerPoint(`.pptx`)** 로 변환하는 **[Agent Skills](https://agentskills.io) 표준 기반 로컬 스킬**입니다.

- **100% 네이티브 객체 변환**: 텍스트, 카드, 배지, 표(`<table>`), Chart.js 차트(`<canvas>`), SVG 다이어그램이 통짜 이미지가 아닌 편집 가능한 PowerPoint 도형·표·차트로 변환됩니다.
- **제로 인프라 & 제로 GCP 설정**: Cloud Run 배포, MCP 서버 데몬, GCP Project ID(`gcloud auth`) 설정 없이 로컬에서 즉시 동작합니다.
- **호스트 에이전트 AI 활용**: 이미지(스크린샷)를 슬라이드로 변환하거나 시각적 품질을 검수할 때는 스킬이 등록된 **호스트 에이전트(Antigravity / Gemini 등) 자신의 멀티모달 AI**를 활용하고, 실제 `.pptx` 생성은 로컬 `scripts/` 엔진이 결정론적으로 수행합니다.

---

## 디렉토리 구조 (`enterprise_skills` 표준)

```text
html-to-pptx/
├── SKILL.md                                  # YAML 프론트매터 + 에이전트 워크플로우 지침
├── scripts/                                  # 로컬 실행 스크립트 및 변환 모듈
│   ├── convert_html_to_pptx.py               # CLI: HTML -> 16:9 네이티브 PPTX 변환
│   ├── browser_renderer.py                   # 헤드리스 Chrome CDP 라이브 DOM 기하 추출 엔진
│   ├── js_geometry_extractor.py              # 브라우저 라이브 DOM 기하·스타일 추출기
│   ├── pptx_native_builders.py               # python-pptx 도형·표·차트 빌더 및 충돌 해결기
│   └── color_utils.py                        # CSS 변수 및 RGBA 알파 블렌딩 파서
├── references/                               # 상세 레퍼런스 문서 (Progressive Disclosure)
│   ├── html_slide_authoring_guide.md         # 이미지->HTML 변환 및 슬라이드 작성용 16:9 시맨틱 가이드
│   └── cli_and_architecture.md               # CLI 플래그, JSON 출력 명세, 커스텀 빌더 훅 설명
└── assets/                                   # 정적 리소스
    ├── fonts/                                # 폴백 폰트 (NotoSansKR 등)
    ├── chart.min.js                          # 오프라인 Chart.js 번들
    └── tailwindcss.min.js                    # 오프라인 Tailwind CSS 번들
```

---

## 설치 방법 (로컬 단독 설치)

Python 3.11+ 및 Chrome(또는 Chromium)만 설치되어 있으면 됩니다. 별도의 GCP 로그인이나 프로젝트 설정이 필요하지 않습니다.

```bash
git clone https://github.com/hwangju1116/html_to_slide.git
cd html_to_slide
./install_antigravity.sh
```

- 가상환경(`.venv`) 생성 및 패키지(`python-pptx`, `lxml`, `websockets`, `pillow`) 설치
- Pretendard 폰트 로컬 캐시
- `~/.gemini/config/skills/html-to-pptx` 심볼릭 링크 자동 등록

---

## 사용 방법

### 1. 에이전트 채팅창에서 자연어로 요청
- **HTML 슬라이드를 PPTX로 변환할 때**
  > "위에서 만든 HTML 슬라이드를 PPTX로 변환해줘"
  > "`~/slides/deck.html`을 PPTX로 변환해줘"
  > "3번 슬라이드만 PPTX로 다시 변환해줘"
- **슬라이드 이미지(`.png` / `.jpg`)를 편집 가능한 PPTX로 변환할 때**
  > "첨부한 슬라이드 이미지를 편집 가능한 PPTX로 만들어줘" *(호스트 AI가 이미지를 분석해 16:9 시맨틱 HTML을 생성한 뒤 네이티브 PPTX로 변환)*

### 2. CLI로 직접 실행

```bash
# 전체 슬라이드를 네이티브 PPTX로 변환
./.venv/bin/python scripts/convert_html_to_pptx.py deck.html -o deck.pptx

# 특정 슬라이드(예: 2번 슬라이드)만 변환 + JSON 결과 출력
./.venv/bin/python scripts/convert_html_to_pptx.py deck.html -s 2 -o slide2.pptx --json
```
