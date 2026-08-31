FROM python:3.12-slim

# OCR (Tesseract + Türkçe dil paketi) ve PDF/görüntü bağımlılıkları
RUN apt-get update && apt-get install -y --no-install-recommends \
        tesseract-ocr \
        tesseract-ocr-tur \
        tesseract-ocr-eng \
        libgl1 \
        poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /srv
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app

ENV FIS_DATA=/data
VOLUME ["/data"]
EXPOSE 8091
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8091"]
