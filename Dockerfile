# json-convert — Extracción de Adobe Analytics con Playwright (Chromium + Firefox)
# Uso: docker build -t ghcr.io/luisalbertomk/json-convert:latest . && docker run --rm ghcr.io/luisalbertomk/json-convert:latest --help

FROM python:3.12-slim

WORKDIR /app

# Instalar dependencias del sistema para Playwright
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    ca-certificates \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libglib2.0-0 \
    libnspr4 \
    libnss3 \
    libu2f-udev \
    libvulkan1 \
    libxcomposite1 \
    libxdamage1 \
    libxfixes3 \
    libxkbcommon0 \
    libxrandr2 \
    xdg-utils \
    && rm -rf /var/lib/apt/lists/*

# Copiar e instalar dependencias Python
COPY config/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt && \
    playwright install chromium firefox && \
    playwright install-deps chromium firefox

# Copiar código
COPY . .

# Entry point
ENTRYPOINT ["python", "src/extract_browser.py"]
