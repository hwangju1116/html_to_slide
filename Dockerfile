FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8080 \
    HOST=0.0.0.0 \
    MCP_TRANSPORT=streamable-http \
    CHROME_BIN=/usr/bin/chromium \
    GOOGLE_GENAI_USE_VERTEXAI=True \
    GOOGLE_CLOUD_LOCATION=global

# Install headless Chromium, CJK/Korean fonts (Noto Sans KR, Nanum), Color Emoji, and Monospace fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    curl \
    fontconfig \
    fonts-noto-cjk \
    fonts-nanum \
    fonts-noto-color-emoji \
    fonts-dejavu-core \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install official Pretendard font weights during image build (zero Git repo bloat)
RUN mkdir -p /usr/local/share/fonts/pretendard && \
    for w in Regular Medium SemiBold Bold ExtraBold Black; do \
      curl -fsSL "https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/packages/pretendard/dist/public/static/Pretendard-${w}.otf" \
        -o "/usr/local/share/fonts/pretendard/Pretendard-${w}.otf" || true; \
    done && \
    fc-cache -fv

WORKDIR /app

COPY pyproject.toml README.md plugin.json ./
COPY assets/ ./assets/
COPY skills/ ./skills/
COPY app/ ./app/

RUN pip install --no-cache-dir .

EXPOSE 8080

CMD ["python", "-m", "app.mcp_server", "--transport", "streamable-http"]
