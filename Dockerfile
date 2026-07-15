FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8000

RUN apt-get update \
    && apt-get install --no-install-recommends -y tesseract-ocr tesseract-ocr-eng tesseract-ocr-hin \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

RUN useradd --create-home --shell /usr/sbin/nologin appuser

COPY requirements.txt ./
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY --chown=appuser:appuser . .

RUN mkdir -p /app/data /app/uploads \
    && chown -R appuser:appuser /app/data /app/uploads

USER appuser

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import os, urllib.request; urllib.request.urlopen('http://127.0.0.1:' + os.getenv('PORT', '8000') + '/', timeout=3)" || exit 1

CMD ["sh", "-c", "exec gunicorn --workers 2 --threads 4 --timeout 120 --bind 0.0.0.0:${PORT:-8000} app:app"]
