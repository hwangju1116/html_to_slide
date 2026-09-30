provider "google" {
  project = var.project_id
  region  = var.region
}

# 1. Enable required GCP APIs in the customer's project
locals {
  required_apis = [
    "run.googleapis.com",
    "aiplatform.googleapis.com",
    "artifactregistry.googleapis.com",
    "cloudbuild.googleapis.com",
    "iam.googleapis.com",
    "cloudresourcemanager.googleapis.com",
  ]
}

resource "google_project_service" "enabled_apis" {
  for_each           = toset(local.required_apis)
  project            = var.project_id
  service            = each.value
  disable_on_destroy = false
}

data "google_project" "current" {
  project_id = var.project_id
  depends_on = [google_project_service.enabled_apis]
}

# 2. Create Artifact Registry Docker repository
resource "google_artifact_registry_repository" "mcp_repo" {
  project       = var.project_id
  location      = var.region
  repository_id = var.service_name
  description   = "Container repository for HTML-to-PPTX MCP Server"
  format        = "DOCKER"

  depends_on = [google_project_service.enabled_apis]
}

# 3. Create dedicated Service Account with Vertex AI User permission
resource "google_service_account" "mcp_sa" {
  project      = var.project_id
  account_id   = "${var.service_name}-sa"
  display_name = "HTML-to-PPTX MCP Server Service Account"

  depends_on = [google_project_service.enabled_apis]
}

resource "google_project_iam_member" "vertex_ai_user" {
  project = var.project_id
  role    = "roles/aiplatform.user"
  member  = "serviceAccount:${google_service_account.mcp_sa.email}"
}

# 4. Build and push container image via Cloud Build (no local Docker daemon required)
locals {
  source_hash = sha256(join("", concat(
    [
      filesha256("${path.module}/../Dockerfile"),
      filesha256("${path.module}/../pyproject.toml"),
    ],
    [for f in sort(fileset("${path.module}/../app", "*.py")) : filesha256("${path.module}/../app/${f}")]
  )))
  image_uri              = "${var.region}-docker.pkg.dev/${var.project_id}/${google_artifact_registry_repository.mcp_repo.repository_id}/${var.service_name}:${substr(local.source_hash, 0, 12)}"
  deterministic_base_url = "https://${var.service_name}-${data.google_project.current.number}.${var.region}.run.app"
}

resource "terraform_data" "cloud_build_image" {
  triggers_replace = [
    local.image_uri,
    local.source_hash,
  ]

  provisioner "local-exec" {
    working_dir = "${path.module}/.."
    command     = "gcloud builds submit --project=${var.project_id} --tag=${local.image_uri} ."
  }

  depends_on = [google_artifact_registry_repository.mcp_repo]
}

# 5. Deploy Cloud Run v2 Service
resource "google_cloud_run_v2_service" "mcp_server" {
  project             = var.project_id
  name                = var.service_name
  location            = var.region
  ingress             = "INGRESS_TRAFFIC_ALL"
  deletion_protection = false

  template {
    service_account = google_service_account.mcp_sa.email
    timeout         = "300s"

    scaling {
      min_instance_count = 0
      max_instance_count = 5
    }

    containers {
      image = local.image_uri

      ports {
        container_port = 8080
      }

      resources {
        limits = {
          cpu    = "2"
          memory = "2Gi"
        }
      }

      env {
        name  = "GOOGLE_CLOUD_PROJECT"
        value = var.project_id
      }
      env {
        name  = "GOOGLE_CLOUD_LOCATION"
        value = var.vertex_location
      }
      env {
        name  = "GOOGLE_GENAI_USE_VERTEXAI"
        value = "True"
      }
      env {
        name  = "MCP_TRANSPORT"
        value = "sse"
      }
      env {
        name  = "PUBLIC_BASE_URL"
        value = local.deterministic_base_url
      }
    }
  }

  depends_on = [
    terraform_data.cloud_build_image,
    google_project_iam_member.vertex_ai_user,
  ]
}

# 6. Optional public invoker binding for SSE MCP clients
resource "google_cloud_run_v2_service_iam_member" "public_invoker" {
  count    = var.allow_unauthenticated ? 1 : 0
  project  = var.project_id
  location = var.region
  name     = google_cloud_run_v2_service.mcp_server.name
  role     = "roles/run.invoker"
  member   = "allUsers"
}

# 7. Automatically register the deployed Cloud Run SSE URL & Skill into local ~/.gemini/config
resource "terraform_data" "register_local_antigravity" {
  count = var.register_local_antigravity ? 1 : 0

  triggers_replace = [
    google_cloud_run_v2_service.mcp_server.uri,
  ]

  provisioner "local-exec" {
    working_dir = "${path.module}/.."
    command     = "bash ./install_antigravity.sh --remote-url ${google_cloud_run_v2_service.mcp_server.uri}/sse"
  }

  depends_on = [google_cloud_run_v2_service.mcp_server]
}
