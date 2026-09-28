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



