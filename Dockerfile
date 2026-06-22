FROM python:3.11-slim

# Installer LibreOffice + dépendances système
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    fonts-dejavu \
    fonts-liberation \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copier et installer les dépendances Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copier le code
COPY . .

# Créer les dossiers nécessaires
RUN mkdir -p generated uploads logs static/css static/js

# Port exposé
EXPOSE 10000

# Lancer gunicorn
CMD gunicorn app:app --bind 0.0.0.0:${PORT:-10000} --workers 2 --timeout 120
