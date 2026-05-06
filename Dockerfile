FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir fastapi "uvicorn[standard]" python-multipart

# Install Playwright Chromium for the downloader
RUN playwright install chromium --with-deps

COPY src/ ./src/
COPY api/ ./api/
COPY static/ ./static/
COPY prompts/ ./prompts/
COPY config.py main.py app.py ./

RUN mkdir -p /app/output /app/data/raw /app/data/processed

EXPOSE 8080

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8080"]
