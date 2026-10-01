FROM python:3.11-slim

# LibreOffice converts the edited .docx back to PDF; fonts keep the layout close to the original
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer fonts-liberation fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]
