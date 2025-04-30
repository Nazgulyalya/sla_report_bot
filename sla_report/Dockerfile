FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    gcc \
    libpq-dev \
    build-essential \
    && apt-get clean

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

WORKDIR /app

COPY python/ ./python/

CMD ["python", "python/main.py"]

