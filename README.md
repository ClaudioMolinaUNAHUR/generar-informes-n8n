## 1. Construir y levantar todos los servicios

```bash
docker compose up -d
```

Esto hace todo automático:
1. **Construye** la imagen de la API LibreOffice (tarda unos minutos la primera vez)
2. **Inicia PostgreSQL** (espera a que esté listo)
3. **Inicia n8n** (espera a PostgreSQL)
4. **Inicia la API LibreOffice**

## 2. Verificar que todo funciona

```bash
docker compose ps
```

Deberías ver los 3 servicios con estado `Up`:

| Servicio      | Container          | Puerto          |
|---------------|--------------------|-----------------|
| `postgres`    | `n8n-postgres`     | `5432`          |
| `n8n`         | `n8n`              | `30227` → 5678  |
| `libreoffice` | `n8n-libreoffice`  | `8000`          |

## 3. Acceder

| Servicio | URL                                | Credenciales              |
|----------|------------------------------------|---------------------------|
| **n8n**  | http://localhost:30227/            | user: `n8n` / pass: `n8n` |
| **API**  | http://localhost:8000/health       | (sin autenticación)       |

## 🔗 Comunicación entre servicios

Desde n8n, la API de LibreOffice se llama por:

```
http://libreoffice:8000
```

Ejemplo desde un nodo HTTP Request en n8n:
- URL: `http://libreoffice:8000/generate`
- URL: `http://libreoffice:8000/health`

## Volúmenes

| Ruta en el host            | Montado en                 | Contenido                          |
|----------------------------|----------------------------|------------------------------------|
| `./data`                   | `/data`                    | Plantillas, PDFs, config, charts   |
| `./libreoffice-python`     | `/app`                     | Código fuente de la API (en vivo)  |
| `postgres-data` (volumen)  | `/var/lib/postgresql/data` | Base de datos PostgreSQL           |
| `libreoffice-tmp` (volumen)| `/tmp`                     | Archivos temporales                |