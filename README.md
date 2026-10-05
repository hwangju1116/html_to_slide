# HTML → PPTX 변환기 (`html-to-pptx`)

HTML 슬라이드를 **편집 가능한 16:9 PowerPoint(.pptx)** 로 변환하는 **Antigravity Skill (및 MCP 서버)** 입니다.
텍스트, 카드, 표, Chart.js 차트, SVG 다이어그램이 통짜 이미지가 아닌 편집 가능한 PowerPoint 도형·표·차트로 변환됩니다.

## 설치 방식 선택

| | **A. Cloud Run (팀 공유)** | **B. 로컬 실행** |
|---|---|---|
| 팀원 PC에 필요한 것 | 없음 (공유받은 명령어 한 줄만 실행) | Python 3.11+, Chrome, gcloud 로그인 |
| 상대 경로 이미지·CSS (`./img/a.png`) | ❌ 반영 안 됨 (외부 URL 또는 인라인 SVG/base64 권장) | ✅ 반영됨 |
| 팀원과 공유 | ✅ URL 하나로 팀 전체 공유 | ❌ 각자 환경 구성 필요 |
| 비용 청구 | 배포한 프로젝트 (Cloud Run + Gemini) | 각자 프로젝트 (Gemini) |

## 사전 준비

```bash
git clone https://github.com/hwangju1116/html_to_slide.git
cd html_to_slide
```

Cloud Run 배포자(관리자)와 로컬 실행 사용자는 결제가 연결된 GCP 프로젝트와 [gcloud](https://cloud.google.com/sdk/docs/install) 로그인이 필요합니다 (Cloud Run 배포 시 [Terraform](https://developer.hashicorp.com/terraform/install) 추가 필요). 이미 배포된 Cloud Run URL을 받아 쓰는 팀원은 아래 로그인 과정을 건너뛰어도 됩니다.

```bash
gcloud auth login
gcloud auth application-default login
```

---

## A. Cloud Run으로 팀 공유하기

### 1. 서버 배포 (관리자 1명)

```bash
terraform -chdir=terraform init
terraform -chdir=terraform apply -var="project_id=YOUR_PROJECT_ID"
```

첫 배포는 10~15분 정도 걸립니다. 배포한 관리자 PC에는 Antigravity Skill과 MCP 서버 등록이 자동으로 완료됩니다. 완료 후 출력되는 `team_install_command`를 팀원에게 공유하세요. (나중에 다시 보려면 `terraform -chdir=terraform output` 실행)

### 2. Antigravity Skill 등록 (팀원 각자)

저장소를 클론한 폴더에서 관리자에게 공유받은 명령어를 실행합니다.

```bash
./install_antigravity.sh --remote-url https://<서비스 URL>/mcp
```

Antigravity Skill과 원격 MCP 서버가 자동으로 연결됩니다. 등록 후 Antigravity에서 `Reload Window`를 실행하세요.

---

## B. 로컬에서 사용하기

```bash
./install_antigravity.sh --project YOUR_PROJECT_ID
```

가상환경 생성, 패키지·폰트 설치, Antigravity Skill 및 MCP 서버 등록을 자동으로 처리합니다. 완료 후 Antigravity에서 `Reload Window`를 실행하세요.

---

## 사용 방법

설치가 완료되면 Antigravity 채팅창에서 자연어로 요청하기만 하면 됩니다. (`html-to-pptx` 스킬이 로컬/Cloud Run 환경에 맞춰 변환부터 `.pptx` 파일 저장까지 자동으로 처리합니다.)

- **같은 대화에서 만든 슬라이드를 바로 변환할 때**
  > "위에서 만든 HTML 슬라이드를 PPTX로 변환해줘"
- **채팅창에 HTML 파일을 첨부하거나 열어둔 상태에서 변환할 때**
  > "첨부한 파일을 PPTX로 변환해줘"
- **특정 파일 경로를 지정하거나 일부 슬라이드만 변환할 때**
  > "`~/slides/deck.html`을 PPTX로 변환해줘"
  > "3번 슬라이드만 PPTX로 다시 변환해줘"

---

## 업데이트 및 삭제

```bash
git pull origin main
terraform -chdir=terraform apply -var="project_id=YOUR_PROJECT_ID"     # 재배포
terraform -chdir=terraform destroy -var="project_id=YOUR_PROJECT_ID"   # 삭제
```

---

## 주의사항

- 기본 설정(`allow_unauthenticated=true`)에서는 URL을 아는 누구나 호출할 수 있고, 비용은 배포한 프로젝트에 청구됩니다. URL은 팀 내부에만 공유하고, 사용하지 않을 때는 삭제하세요.
