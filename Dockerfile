FROM python:3.10-slim

RUN apt-get update && apt-get install -y \
    libxml2 \
    libxslt1-dev \
    python3-dev \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY scripts /app/scripts
WORKDIR /app/scripts

CMD ["python", "main.py"]
