FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    libreoffice-writer \
    libreoffice-common \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir \
    python-pptx \
    pypdf \
    pillow \
    matplotlib \
    numpy \
    fastapi \
    uvicorn \
    python-multipart

# Crear usuario 'n8n_user' y preparar directorios
# Usamos UID 1000 para facilitar la compatibilidad de permisos con volúmenes de n8n en k3s
RUN useradd -m -u 1000 n8n_user && \
    mkdir -p /data && \
    chown -R n8n_user:n8n_user /app /data

# Copiar todo el código modularizado (app.py, services/, utils/, models/)
# Aseguramos que el usuario n8n_user sea el dueño de los archivos
COPY --chown=1000:1000 . .

USER 1000

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
