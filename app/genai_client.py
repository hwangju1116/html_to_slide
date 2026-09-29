import configparser
import os
import subprocess
from typing import Optional
from google import genai

DEFAULT_GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.8-flash")
DEFAULT_VERTEX_LOCATION = "global"


def configure_vertex_environment() -> None:
    """Configures environment variables required for Vertex AI client authentication."""
    os.environ["GOOGLE_API_USE_CLIENT_CERTIFICATE"] = "false"
    if "GOOGLE_API_CERTIFICATE_CONFIG" in os.environ:
        del os.environ["GOOGLE_API_CERTIFICATE_CONFIG"]


def resolve_gcp_project(explicit_project: Optional[str] = None) -> str:
    """Resolves the active GCP project ID from explicit arg, environment, or local gcloud config."""
    if explicit_project and explicit_project.strip():
        return explicit_project.strip()

    for env_key in ("GOOGLE_CLOUD_PROJECT", "GCLOUD_PROJECT", "GCP_PROJECT"):
        val = os.environ.get(env_key, "").strip()
        if val and val != "(unset)":
            return val

    try:
        gcloud_dir = os.environ.get(
            "CLOUDSDK_CONFIG", os.path.expanduser("~/.config/gcloud")
        )
        active_cfg_name = "default"
        active_cfg_file = os.path.join(gcloud_dir, "active_config")
        if os.path.isfile(active_cfg_file):
            with open(active_cfg_file, "r", encoding="utf-8") as f:
                name = f.read().strip()
                if name:
                    active_cfg_name = name
        cfg_path = os.path.join(
            gcloud_dir, "configurations", f"config_{active_cfg_name}"
        )
        if os.path.isfile(cfg_path):
            parser = configparser.ConfigParser()
            parser.read(cfg_path, encoding="utf-8")
            if parser.has_option("core", "project"):
                proj = parser.get("core", "project").strip()
                if proj and proj != "(unset)":
                    os.environ["GOOGLE_CLOUD_PROJECT"] = proj
                    return proj
    except Exception:
        pass

    try:
        out = subprocess.check_output(
            ["gcloud", "config", "get-value", "project"],
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=3,
        ).strip()
        if out and out != "(unset)":
            os.environ["GOOGLE_CLOUD_PROJECT"] = out
            return out
    except Exception:
        pass

    return ""


def verify_gcp_auth(explicit_project: Optional[str] = None) -> tuple[bool, str, str]:
    """Verifies that the caller has an authenticated GCP account (ADC or gcloud) and a valid GCP project ID.

    Returns:
        (is_valid, project_id, error_message)
    """
    configure_vertex_environment()
    project_id = resolve_gcp_project(explicit_project)
    if not project_id:
        return (
            False,
            "",
            (
                "GCP Project ID가 설정되어 있지 않습니다. 이 공유 스킬은 GCP 계정이 있는 사용자만 사용할 수 있습니다.\n"
                "터미널에서 아래 명령어를 실행한 뒤 다시 시도해 주세요:\n"
                "  1) gcloud auth login\n"
                "  2) gcloud auth application-default login\n"
                "  3) ./install_antigravity.sh --project YOUR_GCP_PROJECT_ID"
            ),
        )

    # 1. Cloud Run environment (runs under dedicated GCP Service Account)
    if os.environ.get("K_SERVICE"):
        return True, project_id, ""

    # 2. Check Application Default Credentials file or google.auth
    gcloud_dir = os.environ.get(
        "CLOUDSDK_CONFIG", os.path.expanduser("~/.config/gcloud")
    )
    adc_file = os.environ.get(
        "GOOGLE_APPLICATION_CREDENTIALS",
        os.path.join(gcloud_dir, "application_default_credentials.json"),
    )
    if os.path.isfile(adc_file):
        return True, project_id, ""

    try:
        import google.auth

        creds, _ = google.auth.default(
            scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
        if creds is not None:
            return True, project_id, ""
    except Exception:
        pass

    # 3. Fallback: check active gcloud CLI authenticated account
    try:
        account = subprocess.check_output(
            [
                "gcloud",
                "auth",
                "list",
                "--filter=status:ACTIVE",
                "--format=value(account)",
            ],
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=3,
        ).strip()
        if account:
            return True, project_id, ""
    except Exception:
        pass

    return (
        False,
        project_id,
        (
            f"인증된 GCP 계정을 찾을 수 없습니다 (Project: {project_id}).\n"
            "이 공유 스킬은 GCP 계정 인증이 완료된 사용자만 사용할 수 있습니다. 터미널에서 아래 명령어를 실행해 주세요:\n"
            "  gcloud auth login\n"
            "  gcloud auth application-default login"
        ),
    )


def get_genai_client(
    project: Optional[str] = None,
    location: Optional[str] = None,
) -> genai.Client:
    """Returns an authenticated GenAI client configured for Vertex AI (region: global).

    Args:
        project: GCP Project ID. Defaults to GOOGLE_CLOUD_PROJECT env var or active gcloud project.
        location: Vertex AI Location for Gemini models. Defaults to 'global'.

    Returns:
        genai.Client instance.
    """
    configure_vertex_environment()
    target_project = resolve_gcp_project(project)
    target_location = location or os.environ.get(
        "VERTEX_AI_LOCATION", DEFAULT_VERTEX_LOCATION
    )
    os.environ["GOOGLE_CLOUD_LOCATION"] = target_location

    try:
        kwargs = {"vertexai": True, "location": target_location}
        if target_project:
            kwargs["project"] = target_project
        return genai.Client(**kwargs)
    except Exception:
        return genai.Client()




