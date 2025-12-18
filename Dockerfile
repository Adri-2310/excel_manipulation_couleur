FROM python:3.11-slim

WORKDIR /app

# Dépendances système minimales
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# S'assurer que les dossiers existent
RUN mkdir -p uploads log

EXPOSE 5000

# Lancer l'app Flask via gunicorn (objet app dans app.py) [1]
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]