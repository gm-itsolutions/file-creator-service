FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ ./src/

# Ordner für generierte Dateien
RUN mkdir -p /app/files

ENV FILES_DIR=/app/files
ENV PORT=8002
ENV HOST=0.0.0.0
ENV PYTHONUNBUFFERED=1
# BASE_URL wird in docker-compose gesetzt

EXPOSE 8002

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:8002/health || exit 1

CMD ["python", "src/server.py"]
