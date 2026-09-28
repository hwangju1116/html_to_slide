variable "project_id" {
  description = "Customer/Target GCP Project ID where Cloud Run and Vertex AI will be provisioned."
  type        = string
}

variable "region" {
  description = "GCP region for Cloud Run and Artifact Registry."
  type        = string
  default     = "asia-northeast3"
}

variable "vertex_location" {
  description = "Vertex AI location for Gemini 3.8 Flash calls."
  type        = string
  default     = "global"
}

variable "service_name" {
  description = "Name of the Cloud Run MCP server service."
  type        = string
  default     = "html-to-pptx-mcp"
}

variable "allow_unauthenticated" {
  description = "Whether to allow unauthenticated SSE connections to the Cloud Run MCP endpoint."
  type        = bool
  default     = true
}

variable "register_local_antigravity" {
  description = "Automatically register the deployed Cloud Run MCP server URL and SKILL.md into local ~/.gemini/config after apply."
  type        = bool
  default     = true
}
