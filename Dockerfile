FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1

WORKDIR /app

# Playwright's bundled Node runtime needs libatomic before install-deps can run.
RUN apt-get update && \
    apt-get install -y --no-install-recommends libatomic1 && \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

ENV PLAYWRIGHT_BROWSERS_PATH=/app/pw-browsers
RUN python -m playwright install-deps chromium && \
    python -m playwright install chromium

COPY . .

CMD ["python", "-u", "miner.py"]
