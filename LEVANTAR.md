# 🚀 Instructivo para levantar el proyecto

## 1. Abrir terminal

Pararse en la raíz del proyecto:

```cmd
cd C:\Users\cmolina\Desktop\varios\k8s-n8n
```

## 2. Levantar todo

```cmd
docker compose up -d
```

Esto inicia 3 servicios:
- **PostgreSQL** — base de datos de n8n
- **n8n** — el servidor de automatización de workflows
- **LibreOffice API** — la API que genera PDFs, gráficos, etc.

La primera vez tarda unos minutos porque construye la imagen de la API.

## 3. Verificar que los servicios están corriendo

```cmd
docker compose ps
```

Debe mostrar los 3 contenedores con estado `Up`:

| Contenedor      | Puerto expuesto |
|-----------------|-----------------|
| `n8n-postgres`  | `5432`          |
| `n8n`           | `30227`         |
| `n8n-libreoffice` | `8000`        |

## 4. Acceder a n8n

Abrir en el navegador:

**http://localhost:30227/**

Usuario: `n8n`
Contraseña: `n8n`

## 5. Verificar la API de LibreOffice

Abrir en el navegador o hacer:

```cmd
curl http://localhost:8000/health
```

Debería responder: `{"status": "ok"}`

## 6. Ver los logs en vivo

```cmd
docker compose logs -f
```

Para ver solo un servicio:

```cmd
docker compose logs -f n8n
docker compose logs -f libreoffice
docker compose logs -f postgres
```

## 7. Detener todo

```cmd
docker compose down
```

Los datos de la base de datos y los archivos en `data/` se conservan.

## 8. Para volver a levantar después de detener

```cmd
docker compose up -d
```

(Sin necesidad de rebuild, a menos que haya cambios en el Dockerfile)

## 9. Reconstruir la imagen de la API (si se modificó el Dockerfile)

```cmd
docker compose build
docker compose up -d
```

---

## 🔗 Cómo se comunican n8n y la API

Dentro de docker-compose, los servicios se ven entre sí por nombre.  
Desde n8n, la API de LibreOffice se llama con:

```
http://libreoffice:8000
```

Ejemplos de endpoints:
- `http://libreoffice:8000/health`
- `http://libreoffice:8000/generate`
- `http://libreoffice:8000/generate-grafs`

---

## 📁 Archivos compartidos

| Carpeta local    | Se monta en      | Para qué                          |
|------------------|------------------|-----------------------------------|
| `./data`         | `/data`          | Plantillas, PDFs, configuraciones |
| `./libreoffice-python` | `/app`     | Código de la API (cambios en vivo) |