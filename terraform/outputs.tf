output "cloud_run_service_url" {
  description = "Base URL of the deployed Cloud Run MCP service."
  value       = google_cloud_run_v2_service.mcp_server.uri
}

output "mcp_sse_url" {
  description = "MCP SSE endpoint URL to register in ~/.gemini/config/mcp_config.json."
  value       = "${google_cloud_run_v2_service.mcp_server.uri}/sse"
}

output "team_install_command" {
  description = "One-line command for team members to connect their Antigravity globally to this Cloud Run MCP server."
  value       = "./install_antigravity.sh --remote-url ${google_cloud_run_v2_service.mcp_server.uri}/sse"
}

output "mcp_config_json_snippet" {
  description = "JSON snippet for ~/.gemini/config/mcp_config.json"
  value = jsonencode({
    mcpServers = {
      "html-to-pptx" = {
        serverUrl = "${google_cloud_run_v2_service.mcp_server.uri}/sse"
      }
    }
  })
}
