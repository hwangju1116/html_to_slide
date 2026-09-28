FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080 \
    HOST=0.0.0.0 \
    MCP_TRANSPORT=sse \
    CHROME_BIN=/usr/bin/chromium \
    GOOGLE_GENAI_USE_VERTEXAI=True \
    GOOGLE_CLOUD_LOCATION=global

# Install headless Chromium and CJK font support
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fontconfig \
    fonts-noto-cjk \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy bundled Pretendard fonts into system font cache for pixel-perfect Korean slide rendering
COPY assets/fonts/ /usr/local/share/fonts/pretendard/
RUN fc-cache -fv

COPY pyproject.toml README.md plugin.json ./
COPY assets/ ./assets/
COPY skills/ ./skills/
COPY app/ ./app/

RUN pip install --no-cache-dir .

EXPOSE 8080

CMD ["python", "-m", "app.mcp_server", "--transport", "sse"]

