# HTML-to-PPTX MCP Server & Antigravity Skill (`html-to-pptx-mcp`)

Standalone MCP Server and Antigravity Skill package that converts single-slide or multi-slide HTML presentations into **100% editable native 16:9 PowerPoint (`.pptx`)** decks.

Supports **both** **Remote Cloud Run (`SSE`) deployment** (automated via Terraform into any customer/team GCP project) and **Local (`stdio`) execution** (with automatic `gcloud` project detection).

## How It Works

Instead of pasting static screenshots, `html-to-pptx-mcp` converts HTML slides into **editable PowerPoint elements**:

1. **Browser Rendering**: Opens the HTML in a 16:9 headless browser to apply fonts, CSS layouts, and charts.
2. **Layout & Style Extraction**: Reads the position, colors, tables, and chart data of every element on the slide.
3. **Native PPTX Generation**: Converts cards into PPT shapes, `<canvas>` charts into PPT charts, `<table>` into PPT tables, and text into editable text boxes.
4. **Layout Polish**: Automatically adjusts layer order and spacing to prevent overlapping elements.

---

## Prerequisites: GCP Authentication

Before deploying via Terraform or running locally with Vertex AI, authenticate your `gcloud` CLI and Application Default Credentials (ADC):

```bash
# 1. Authenticate Application Default Credentials (used by Terraform & Vertex AI SDK)
gcloud auth application-default login

# 2. Authenticate gcloud CLI session (used by Cloud Build during Terraform deployment)
gcloud auth login
```

---

## Option A: Automated Cloud Run Deployment via Terraform (Recommended for Customers & Teams)

Each customer or team can provision an isolated **Cloud Run MCP Server** inside **their own GCP project** (preventing cross-tenant billing or slide data sharing) with a single Terraform command.

### 1. Provision Cloud Run + Vertex AI IAM + Artifact Registry
```bash
cd terraform
terraform init
terraform apply -var="project_id=YOUR_GCP_PROJECT_ID"
```
What Terraform does automatically:
1. Enables required APIs (`run.googleapis.com`, `aiplatform.googleapis.com`, `artifactregistry.googleapis.com`, `cloudbuild.googleapis.com`, `iam.googleapis.com`).
2. Creates a dedicated Service Account (`html-to-pptx-mcp-sa`) with `roles/aiplatform.user`.
3. Builds the container image remotely via **Cloud Build** (no local Docker required) with Headless Chromium and Pretendard fonts pre-installed.
4. Deploys the Cloud Run SSE service (`https://html-to-pptx-mcp-<project_num>.<region>.run.app/sse`) and registers it in your local `~/.gemini/config/mcp_config.json`.

### 2. Share with Team Members (Zero Local Setup)
Once deployed to Cloud Run, other team members do **not** need Python, Chromium, or `gcloud` configured locally. They can attach the Skill and remote MCP server globally with one command:
```bash
./install_antigravity.sh --remote-url https://html-to-pptx-mcp-xxxxx.asia-northeast3.run.app/sse
```
Or by adding `serverUrl` directly to `~/.gemini/config/mcp_config.json`:
```json
{
  "mcpServers": {
    "html-to-pptx": {
      "serverUrl": "https://html-to-pptx-mcp-xxxxx.asia-northeast3.run.app/sse"
    }
  }
}
```

---

## Option B: Local Installation (Auto-Detects User's `gcloud` Project)

If running locally over `stdio`, `install_antigravity.sh` automatically detects your active `gcloud config get-value project` (no manual config editing needed), creates `.venv` if missing, locks the Vertex AI model location to `global`, and links the Skill & MCP server into `~/.gemini/config/`:

```bash
chmod +x install_antigravity.sh
./install_antigravity.sh --project YOUR_GCP_PROJECT_ID
```
*(If `--project` is omitted, `install_antigravity.sh` automatically uses `gcloud config get-value project`.)*

---

## Troubleshooting

### 1. `oauth2: "invalid_grant" "reauth related error (invalid_rapt)"` during `terraform apply`
- **Cause**: Your Google Cloud Application Default Credentials (ADC) or `gcloud` session token has expired (RAPT re-authentication policy).
- **Fix**: Refresh both credentials in your terminal and re-run `terraform apply`:
  ```bash
  gcloud auth application-default login
  gcloud auth login
  terraform apply -var="project_id=YOUR_GCP_PROJECT_ID"
  ```

### 2. `PERMISSION_DENIED: Permission 'businessaicode.locations.queryConfiguration' denied` in Antigravity
- **Cause**: Antigravity IDE itself picked up your global `gcloud` default project, which does not have the Gemini for Google Cloud (`cloudaicompanion.googleapis.com`) API enabled.
- **Fix**: Unset the global `gcloud` default project so Antigravity uses its default login, and pass `--project` explicitly to `install_antigravity.sh` so only the `html-to-pptx` MCP server uses that GCP project for Vertex AI:
  ```bash
  gcloud config unset project
  ./install_antigravity.sh --project YOUR_GCP_PROJECT_ID
  ```


