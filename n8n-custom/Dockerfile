FROM n8nio/n8n:1.74.0

USER root

RUN apk add --no-cache \
        python3 \
        py3-pip \
        libreoffice \
        libreoffice-common \
        libreoffice-writer \
        libreoffice-calc \
        libreoffice-impress \
        ttf-dejavu \
    && pip install --break-system-packages \
        python-pptx \
        pypdf \
        pillow \
        matplotlib \
        requests

USER node