FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-impress \
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
    uvicorn

# Crear usuario 'n8n' (UID 1000) para coincidir con el usuario oficial de n8n
RUN useradd -m -u 1000 n8n_user
USER 1000

COPY --chown=1000:1000 app.py .

EXPOSE 8000

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
