FROM python:3.12-slim

# Zapewnia możliwość instalacji pakietów wymagających kompilacji
RUN apt-get update && apt-get install -y \
    build-essential \
    libxml2-dev \
    libxslt1-dev \
    libffi-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .

# Instalujemy bez zbędnych cachy
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY word_extractor.py .
COPY batch_runner.py .

# Tworzymy katalogi, o ile nie istnieją
RUN mkdir -p /app/input_files /app/output_files

CMD ["python", "batch_runner.py"]
