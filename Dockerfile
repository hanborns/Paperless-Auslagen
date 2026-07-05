FROM python:3.12-slim

# poppler-utils wird von pdf2image für die PDF->PNG-Konvertierung benötigt
RUN apt-get update \
    && apt-get install -y --no-install-recommends poppler-utils \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
    flask==3.* \
    requests==2.* \
    openpyxl==3.* \
    pdf2image==1.* \
    pillow==11.*

WORKDIR /app
COPY worker.py .

EXPOSE 5000
CMD ["python", "worker.py"]
