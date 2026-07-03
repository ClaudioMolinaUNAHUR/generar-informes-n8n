FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    git \
    libreoffice-impress \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-common \
    fonts-dejavu \
    fonts-liberation \
    procps \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
    python-pptx \
    openpyxl \
    pypdf \
    pillow \
    matplotlib \
    numpy \
    fastapi \
    uvicorn \
    python-multipart \
    pydantic \
    openpyxl

# Crear usuario 'n8n_user' y preparar directorios
RUN useradd -m -u 1000 n8n_user && \
    mkdir -p /data && \
    chown -R n8n_user:n8n_user /app /data

# Copiamos primero los requerimientos o el código si no usamos volúmenes
# Para el flujo de "git pull" externo, esto se sobreescribirá con el montaje
COPY --chown=n8n_user:n8n_user . .

USER 1000

EXPOSE 8000

# Usamos --reload si estamos en modo desarrollo para que detecte el git pull
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000", "--reload"]