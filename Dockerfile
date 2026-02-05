FROM python:3.11-slim

WORKDIR /app

# Install system dependencies for reportlab and Pillow
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libjpeg-dev \
    zlib1g-dev \
    libfreetype6-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements first for better caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code
COPY src/ ./src/

# Create directories for files, assets, and templates
RUN mkdir -p /app/files /app/assets/logos /app/assets/images /app/templates/docx /app/templates/xlsx /app/templates/pptx

# Set permissions
RUN chmod -R 755 /app

EXPOSE 8002

CMD ["python", "src/server.py"]
