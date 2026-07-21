# n8n + LibreOffice Python API — Despliegue en k3s (Kubernetes)

## 📋 Visión General

Este repositorio despliega un entorno de **n8n** (workflow automation) con una **API Python/LibreOffice** externa para generación de reportes PDF, gráficos, y manipulación de documentos. Todo orquestado en **k3s (Kubernetes ligero)** dentro de un namespace dedicado `n8n-project`.

El diseño separa la API de procesamiento de documentos (LibreOffice + Python) del contenedor de n8n, permitiendo desarrollar y actualizar el código de la API sin tener que reconstruir imágenes Docker.

---

## 🏗️ Arquitectura

```
┌─────────────────────────────────────────────────────────────┐
│                    k3s Cluster                              │
│  ┌─────────────────────┐  ┌──────────────────────────────┐  │
│  │   Namespace         │  │  Namespace                   │  │
│  │   n8n-project       │  │  ingress-nginx               │  │
│  │                     │  │                              │  │
│  │  ┌───────────────┐  │  │  ┌───────────────────┐       |  │
│  │  │ n8n           │  │  │  │ nginx-ingress     │       │  │
│  │  │ Deployment    │  │  │  │ Controller        │       │  │
│  │  │ port 5678     │  │  │  └───────────────────┘       │  │
│  │  └───────┬───────┘  │  └──────────────────────────────┘  │
│  │          │          │                                    │
│  │  ┌───────▼───────┐  │                                    │
│  │  │ n8n-service   │◄─┤── NodePort :<NODE_PORT> (ej: 30227)│
│  │  │ ClusterIP     │  │   Ingress: / → n8n-service:80      │
│  │  └───────┬───────┘  │                                    │
│  │          │          │                                    │
│  │  ┌───────▼───────┐  │                                    │
│  │  │ PostgreSQL    │  │  ┌──────────────────────────────┐  │
│  │  │ StatefulSet   │  │  │ libreoffice-python           │  │
│  │  │ port 5432     │  │  │ Deployment  port 8000        │  │
│  │  └───────┬───────┘  │  └──────────────────────────────┘  │
│  │          │          │                                    │
│  │  ┌───────▼───────┐  │                                    │
│  │  │ postgres-     │  │  ┌──────────────────────────────┐  │
│  │  │ service       │  │  │ libreoffice-svc              │  │
│  │  │ ClusterIP     │  │  │ ClusterIP → port 80:8000     │  │
│  │  └───────────────┘  │  └──────────────────────────────┘  │
│  │                     │                                    │
│  │  ┌──────────────────────────────────────────────────┐    |
│  │  │ PersistentVolume / PersistentVolumeClaim         │    │
│  │  │                                                  │    │
│  │  │ n8n-pvc → /data  (compartido entre n8n y api)    │    │
│  │  │ libreoffice-code-pvc → /app  (código de la API)  │    │
│  │  │ postgres-data → /var/lib/postgresql/data         │    │
│  │  └──────────────────────────────────────────────────┘    │
│  └────────────────────────────────────────────────────────-─┘
```

### 🔗 Flujo de comunicación

1. **n8n** ejecuta workflows que necesitan generar PDFs, gráficos o manipular documentos.
2. **n8n** se comunica con la API Python mediante **Python Task Runner** (`N8N_PYTHON_TASK_RUNNER_URL`) apuntando a `http://libreoffice-svc.n8n-project.svc.cluster.local:80`.
3. La **API LibreOffice Python** procesa los requests: genera PPTX, convierte a PDF con LibreOffice, crea gráficos con matplotlib, etc.
4. Los archivos generados se guardan en `/data` (**volumen compartido**) para que n8n pueda acceder a ellos.
5. **PostgreSQL** almacena los datos de n8n (workflows, credenciales, ejecuciones).

---

## 🗂️ Estructura del Repositorio

```
k8s-n8n/
├── n8n-k3s/                          # Manifiestos Kubernetes
│   ├── n8n-deployment.yaml           # Deployment de n8n
│   ├── n8n-service.yaml              # Service de n8n (NodePort, actualmente 30227)
│   ├── n8n-configmap.yaml            # Variables de entorno de n8n
│   ├── n8n-secrets.yaml              # Secretos de n8n (DB pass, auth, encryption key)
│   ├── n8n-volume.yaml               # PersistentVolumes + Claims
│   ├── libreoffice-deployment.yaml   # Deployment de la API Python/LibreOffice
│   ├── libreoffice-service.yaml      # Service de la API (ClusterIP)
│   ├── postgres-statefulset.yaml     # StatefulSet de PostgreSQL
│   ├── postgres-service.yaml         # Service Headless de PostgreSQL
│   ├── postgres-secrets.yaml         # Secretos de PostgreSQL
│   ├── ingress/                      # Ingress NGINX
│   │   ├── ingress-n8n-class.yml     # IngressClass para nginx
│   │   ├── n8n-ingress.yml           # Reglas de Ingress
│   │   └── nginx-ingress-n8n-service.yml  # Service del Ingress Controller
│   └── README.md                     # Este archivo
│
├── libreoffice-python/               # Código fuente de la API (montado por PVC)
│   ├── app.py                        # FastAPI principal
│   ├── Dockerfile                    # Docker image definition
│   ├── models/
│   │   └── schemas.py                # Pydantic schemas
│   ├── services/
│   │   ├── pdf_service.py            # Generación de PPTX y conversión a PDF
│   │   ├── structure_service.py      # Construcción de estructura de slides
│   │   ├── graf_service.py           # Generación de gráficos (matplotlib)
│   │   ├── chart_service.py          # Creación de chart images
│   │   ├── xlsx_service.py           # Conversión XLSX → PDF
│   │   └── merge_service.py          # Fusión de PDFs
│   ├── utils/
│   │   └── helpers.py                # Funciones auxiliares (rutas, logos, etc.)
│   └── tests/                        # Tests
│
├── n8n-custom/
│   └── Dockerfile                    # (Obsoleto) Imagen n8n con Python embebido
│
├── data/                             # Volumen de datos compartido (/data)
│   ├── charts/                       # Configuraciones de gráficos por producto
│   ├── config/                       # Archivos Excel de configuración
│   ├── plantillas/                   # Plantillas PPTX para cada producto
│   ├── generados/                    # Reportes PDF generados
│   ├── base-pdf/                     # Plantillas PDF base + imágenes
│   ├── pdf-parts/                    # Partes de PDF generados
│   ├── pptx-parts/                   # Partes de PPTX generados
│   ├── orden-de-pago/                # Archivos de orden de pago
│   └── *.py                          # Scripts de generación
```

---

## 🚀 Despliegue

### 1️⃣ Prerrequisitos

- **k3s** instalado y funcionando (o cualquier cluster Kubernetes 1.19+)
- **kubectl** configurado para apuntar al cluster
- **(Opcional)** Helm si se usa el Ingress NGINX

### 2️⃣ Namespace

Todos los recursos se despliegan en el namespace `n8n-project`:

```bash
kubectl create namespace n8n-project
```

### 3️⃣ Volúmenes Persistentes

> **⚠️ IMPORTANTE**: Los PersistentVolume (PV) en `n8n-volume.yaml` usan `hostPath` con rutas de un entorno de desarrollo (`/mnt/c/Users/fmora/Desktop/...`).  
> **Para producción**, debes modificar estas rutas a ubicaciones válidas en los nodos del cluster o, mejor aún, usar un **StorageClass** apropiado (NFS, Longhorn, Rook/Ceph, etc.).
>
> **⚠️ PERMISOS**: El directorio `/data` debe tener permisos de escritura para el usuario **uid 1000** (n8n_user), ya que tanto n8n como la API guardan, modifican y eliminan archivos constantemente en él. Asegúrate de que la ruta hostPath tenga permisos adecuados:
> ```bash
> # Ejemplo: dar permisos de escritura al uid 1000 en la ruta del host
> chown -R 1000:1000 /ruta/a/k8s-n8n/data
> chmod -R 755 /ruta/a/k8s-n8n/data
> ```

Los PVs definidos son:

| PV | Ruta hostPath (ejemplo) | PVC | Mount en container |
|---|---|---|---|
| `n8n-pv` | `/mnt/c/Users/fmora/Desktop/.../data` | `n8n-pvc` | `/data` (n8n + libreoffice) |
| `libreoffice-code-pv` | `/mnt/c/Users/fmora/Desktop/.../libreoffice-python` | `libreoffice-code-pvc` | `/app` (libreoffice) |
| `postgres-pv` | `/mnt/c/Users/fmora/Desktop/.../postgres-data` | `postgres-data-postgres-statefulset-0` | `/var/lib/postgresql/data` |

```bash
kubectl apply -f n8n-k3s/n8n-volume.yaml
```

### 4️⃣ PostgreSQL

Base de datos de n8n:

```bash
kubectl apply -f n8n-k3s/postgres-secrets.yaml
kubectl apply -f n8n-k3s/postgres-service.yaml
kubectl apply -f n8n-k3s/postgres-statefulset.yaml
```

Verificar que PostgreSQL esté listo:

```bash
kubectl -n n8n-project get pods -l app=postgres
```

### 5️⃣ Configuración de n8n

```bash
kubectl apply -f n8n-k3s/n8n-configmap.yaml
kubectl apply -f n8n-k3s/n8n-secrets.yaml
kubectl apply -f n8n-k3s/n8n-service.yaml
kubectl apply -f n8n-k3s/n8n-deployment.yaml
```

### 6️⃣ API LibreOffice Python

Primero, **cada equipo debe construir su propia imagen Docker** a partir del Dockerfile provisto y pushearla a su propio registry (DockerHub, Harbor, etc.):

```bash
cd libreoffice-python

# Construir la imagen
docker build -t <tu-registry>/api-libreoffice-python:<version> .

# Pushear al registry
docker push <tu-registry>/api-libreoffice-python:<version>
```

Luego, actualizar `libreoffice-deployment.yaml` con la imagen construida:

```yaml
# En libreoffice-deployment.yaml, reemplazar:
image: <tu-registry>/api-libreoffice-python:<version>
```

> **Nota**: La imagen `claudito16/api-libreoffice-python:latest` usada actualmente en el YAML es del desarrollador y solo sirve como referencia. El equipo de infraestructura debe generar su propia imagen a partir del `Dockerfile` en `libreoffice-python/`.

Finalmente, aplicar los manifests:

```bash
kubectl apply -f n8n-k3s/libreoffice-service.yaml
kubectl apply -f n8n-k3s/libreoffice-deployment.yaml
```

### 7️⃣ (Opcional) Ingress NGINX

Si quieres acceso HTTP desde fuera del cluster:

```bash
kubectl apply -f n8n-k3s/ingress/
```

Para usar el Ingress, necesitas tener el **NGINX Ingress Controller** instalado en tu cluster:

```bash
kubectl apply -f https://raw.githubusercontent.com/kubernetes/ingress-nginx/controller-v1.10.0/deploy/static/provider/cloud/deploy.yaml
```

---

## ✅ Verificación del Despliegue

```bash
# Ver todos los recursos
kubectl -n n8n-project get all

# Ver pods en ejecución
kubectl -n n8n-project get pods

# Ver logs de la API
kubectl -n n8n-project logs -l app=libreoffice

# Ver logs de n8n
kubectl -n n8n-project logs -l app=n8n

# Probar health check de la API
kubectl -n n8n-project exec deploy/libreoffice-deployment -- curl -s http://localhost:8000/health

# Probar health check de n8n
kubectl -n n8n-project exec deploy/n8n-deployment -- curl -s http://localhost:5678/healthz
```

---

## 🔌 Acceso a n8n

| Método | URL (ejemplo) | Descripción |
|---|---|---|
| **NodePort** | `http://<NODO_IP>:<NODE_PORT>` | Acceso directo por IP del nodo. El puerto se define en `n8n-service.yaml` (actualmente `nodePort: 30227`, pero puede cambiar según el entorno). |
| **Ingress** | `http://n8n.local/` o tu dominio con HTTPS | Requiere DNS/config local y configuración de TLS/SSL |

> **⚠️ IMPORTANTE**: El puerto `30227` y la URL `http://localhost:30227/` usadas en este README y en los YAML son valores **de ejemplo/desarrollo**.  
> En producción, el equipo de infraestructura puede cambiar:
> - El `nodePort` en `n8n-service.yaml` (o usar `type: ClusterIP` sin NodePort si usan Ingress)
> - El `WEBHOOK_URL` en `n8n-deployment.yaml` para reflejar el dominio real (ej: `https://n8n.miempresa.com/`)
> - El protocolo a **HTTPS** configurando TLS en el Ingress

**Credenciales por defecto:**
- User: `n8n`
- Password: `n8n`

(Pueden cambiarse editando `n8n-secrets.yaml`.)

---

## 🐍 API LibreOffice Python — Endpoints

### Endpoints Principales

| Endpoint | Método | Descripción |
|---|---|---|
| `/health` | GET | Health check de la API |
| `/generate` | POST | Genera un reporte completo (PPTX → PDF) |
| `/generate-grafs` | POST | Genera gráficos de torta como imágenes base64 |
| `/build-structure` | POST | Construye la estructura de slides a partir de datos |
| `/generate-n-emp` | POST | Genera PDF combinado para múltiples empresas |
| `/xlsx-to-pdf` | POST | Convierte un Excel a PDF (rango de celdas) |
| `/merge-pdfs` | POST | Une múltiples PDFs en uno solo |
| `/files/read` | GET | Lee/descarga archivos de `/data` |
| `/files/list` | GET | Lista archivos de `/data` |
| `/files/save` | POST | Guarda archivos en `/data` |

### Comunicación n8n → API

n8n se conecta a la API mediante la variable `N8N_PYTHON_TASK_RUNNER_URL` configurada en `n8n-deployment.yaml`:

```yaml
- name: N8N_PYTHON_TASK_RUNNER_URL
  value: "http://libreoffice-svc.n8n-project.svc.cluster.local:80"
```

Esto permite que n8n ejecute código Python que llama a los endpoints de la API.

---

## 💾 Volúmenes y Datos

### Volumen Compartido `/data`

El directorio `/data` es compartido entre **n8n** y **libreoffice-python** mediante el PVC `n8n-pvc`. Esto permite:

1. **n8n** guarda/lee archivos directamente en `/data` (ej: inputs JSON, plantillas)
2. **API Python** procesa esos archivos y guarda resultados (PDFs generados)
3. Ambos pueden acceder a la misma configuración y templates

> **⚠️ PERMISOS**: La API escribe y elimina archivos constantemente en `/data` (PDFs generados, partes, etc.). El contenedor corre con **uid 1000**, por lo que la ruta hostPath del PV debe tener permisos de escritura para ese usuario:
> ```bash
> # En el nodo donde está montado el PV
> sudo chown -R 1000:1000 /ruta/a/k8s-n8n/data
> sudo chmod -R 755 /ruta/a/k8s-n8n/data
> ```

### Volumen de Código `/app`

El directorio `/app` de la API se monta desde un PV separado (`libreoffice-code-pvc`) que apunta al directorio `libreoffice-python/` del host. Esto permite:

- **Desarrollar y actualizar** el código de la API sin reconstruir la imagen Docker.
- **Reflejar cambios** al instante (el contenedor corre `uvicorn --reload`).
- Hacer `git pull` en el host y los cambios se ven reflejados en el pod.

> **⚠️ En producción**, considera usar un enfoque más robusto:
> - Construir y pushear una nueva imagen Docker a un registry propio (no usar imágenes de terceros)
> - Actualizar la imagen en el deployment (`kubectl set image`)
> - Asegurar que el directorio `/data` en el host tenga permisos de escritura para uid 1000

---

## 🧪 Desarrollo y Actualización de la API

### Flujo de trabajo

```bash
# 1. Clonar el repo en el servidor/host
git clone <repo-url> /ruta/a/k8s-n8n

# 2. Asegurar que el PV apunte a esta ruta (n8n-volume.yaml)
#    libreoffice-code-pv → hostPath: "/ruta/a/k8s-n8n/libreoffice-python"

# 3. Hacer cambios en el código local
#    (agregar endpoints, modificar servicios, etc.)

# 4. Los cambios se ven reflejados automáticamente gracias a --reload

# 5. Para confirmar que funciona, reiniciar el pod
kubectl -n n8n-project delete pod -l app=libreoffice
```

### Construir la imagen Docker (cuando sea necesario)

Cada equipo debe construir su propia imagen a partir del `Dockerfile` en `libreoffice-python/`:

```bash
cd libreoffice-python
docker build -t <tu-registry>/api-libreoffice-python:<version> .
docker push <tu-registry>/api-libreoffice-python:<version>

# Actualizar el deployment con la nueva imagen
kubectl -n n8n-project set image deployment/libreoffice-deployment libreoffice=<tu-registry>/api-libreoffice-python:<version>
# O bien editar el archivo YAML y re-aplicar
```

---

## ⚙️ Variables de Entorno Importantes

### n8n ConfigMap (`n8n-configmap.yaml`)

| Variable | Descripción |
|---|---|
| `NODE_ENV` | Entorno (production) |
| `GENERIC_TIMEZONE` | Zona horaria |
| `WEBHOOK_TUNNEL_URL` | URL para webhooks |
| `DB_TYPE` | Tipo de BD (postgresdb) |
| `DB_POSTGRESDB_USER` | Usuario BD |
| `DB_POSTGRESDB_DATABASE` | Nombre BD |
| `DB_POSTGRESDB_HOST` | Host BD (postgres-service) |
| `DB_POSTGRESDB_PORT` | Puerto BD (5432) |
| `N8N_BASIC_AUTH_ACTIVE` | Autenticación básica activa |
| `N8N_BASIC_AUTH_USER` | Usuario auth básica |

### n8n Secrets (`n8n-secrets.yaml`)

> ⚠️ **IMPORTANTE**: Los valores actuales son para desarrollo. **CAMBIAR las contraseñas en producción**.

| Variable | Valor por defecto | Descripción |
|---|---|---|
| `DB_POSTGRESDB_PASSWORD` | `n8n` | Password BD |
| `N8N_BASIC_AUTH_PASSWORD` | `n8n` | Password auth básica |
| `N8N_ENCRYPTION_KEY` | UUID fijo | Clave de encriptación de datos |

### n8n Deployment (adicionales)

| Variable | Valor (ejemplo) | Descripción |
|---|---|---|
| `WEBHOOK_URL` | `http://localhost:30227/` | URL base de webhooks. **Cambiar según el entorno**: si usan HTTPS con dominio propio, ej: `https://n8n.miempresa.com/` |
| `N8N_PYTHON_TASK_RUNNER_URL` | `http://libreoffice-svc.n8n-project.svc.cluster.local:80` | URL del Task Runner Python (comunicación interna en el cluster, no expuesta) |

---

## 📊 Recursos (Resources)

| Componente | CPU Límite | Memoria Límite | CPU Request | Memoria Request |
|---|---|---|---|---|
| **n8n** | 1.0 core | 1024 Mi | 0.5 core | 512 Mi |
| **libreoffice-python** | 1.0 core | 1024 Mi | 0.5 core | 512 Mi |

Los recursos pueden ajustarse en los archivos de deployment según la carga esperada.

---

## 🔄 Troubleshooting

### La API no responde

```bash
# Verificar el pod
kubectl -n n8n-project describe pod -l app=libreoffice

# Ver logs
kubectl -n n8n-project logs -l app=libreoffice --tail=50

# Verificar el service
kubectl -n n8n-project get svc libreoffice-svc

# Hacer un curl interno
kubectl -n n8n-project run curl-test --image=curlimages/curl -it --rm --restart=Never -- curl -s http://libreoffice-svc:80/health
```

### n8n no se conecta a PostgreSQL

```bash
# Verificar PostgreSQL
kubectl -n n8n-project logs -l app=postgres --tail=20

# Verificar conectividad
kubectl -n n8n-project run pg-test --image=postgres:15-alpine -it --rm --restart=Never -- \
  psql -h postgres-service -U n8n -d n8n -c "SELECT 1"
```

### Los cambios en el código no se reflejan

```bash
# Verificar que el PVC esté montado
kubectl -n n8n-project exec deploy/libreoffice-deployment -- ls -la /app

# Verificar que el PV apunte al directorio correcto
kubectl get pv libreoffice-code-pv -o yaml

# Forzar reinicio del pod
kubectl -n n8n-project rollout restart deployment/libreoffice-deployment
```

### Error de permisos en /data

```bash
# Verificar permisos actuales dentro del pod
kubectl -n n8n-project exec deploy/libreoffice-deployment -- ls -la /data

# Verificar el usuario con el que corre el contenedor
kubectl -n n8n-project exec deploy/libreoffice-deployment -- id

# Solución: corregir permisos en el host (nodo donde está montado el PV)
# ssh al nodo y ejecutar:
sudo chown -R 1000:1000 /ruta/a/k8s-n8n/data
sudo chmod -R 755 /ruta/a/k8s-n8n/data
```

---

## ⚠️ Notas para Producción

1. **Cambiar secrets**: Actualizar todas las contraseñas en `n8n-secrets.yaml` y `postgres-secrets.yaml` antes de desplegar en producción. Usar `base64` o usar un gestor como SealedSecrets / External Secrets Operator.

2. **Storage**: Reemplazar `hostPath` por un storage class adecuado (Longhorn, NFS, Rook/Ceph) para datos persistentes.

3. **Imagen Docker de la API**: La imagen `claudito16/api-libreoffice-python:latest` es del desarrollador. **Cada equipo debe construir su propia imagen** a partir del `Dockerfile` en `libreoffice-python/` y pushearla a su propio registry.

4. **Código de la API**: En producción, considera construir y pushear la imagen Docker completa en lugar de montar el código por volumen. Alternativamente, usa un initContainer que haga `git clone`.

5. **Permisos /data**: Asegurar que el directorio `/data` tenga permisos de escritura para uid 1000, ya que la API escribe y elimina archivos constantemente.

6. **Ingress**: Configurar TLS/SSL en el Ingress para acceso HTTPS.

7. **Backups**: Implementar backups periódicos de PostgreSQL y del directorio `/data`.

8. **Actualizar encryption key**: Si cambia `N8N_ENCRYPTION_KEY`, n8n no podrá desencriptar datos existentes. Guardarla de forma segura.

9. **Replicas**: Para alta disponibilidad, n8n soporta múltiples réplicas (actualmente `replicas: 1`). Requiere compartir el volumen `/data` en modo `ReadWriteMany`.

---

## 🏛️ Diagrama de Componentes

```
          INTERNET
              │
              ▼
     ┌───────────────-------------------─┐
     │ NodePort:<NODE_PORT> (ej: 30227)  │
     │ Ingress: nginx                    │
     └────────┬─────-------------------──┘
              │
      ┌───────▼────────┐
      │   n8n-service  │
      │   ClusterIP    │
      └───────┬────────┘
              │
      ┌───────▼────────────────────┐
      │       n8n (Deployment)     │
      │  ────────────────────────  │
      │  WEBHOOK_URL               │
      │  N8N_PYTHON_TASK_RUNNER_URL│ ────→ http://libreoffice-svc:80
      │  ┌─────────────────────┐   │
      │  │ Python Code Node    │   │
      │  │ (Task Runner)       │───┼───→ llama API Python
      │  └─────────────────────┘   │
      │  Volume: /data (n8n-pvc)   │
      └──────────┬───────────────-─┘
                 │
      ┌──────────▼────────────────┐
      │  PostgreSQL (StatefulSet) │
      │  DB: n8n                  │
      │  User: n8n                │
      │  Volume: postgres-data    │
      └───────────────────────────┘

┌──────────────────────────────────────────────────────────┐
│       libreoffice-python (Deployment)                    │
│  ──────────────────────────────────────────────────────  │
│  Image: <tu-registry>/api-libreoffice-python:<version>   │
│  CMD: uvicorn app:app --reload                           │
│  Port: 8000                                              │
│  ┌──────────────────────────────────────────────────┐    │
│  │ FastAPI Endpoints:                               │    │
│  │  POST /generate        → Reporte completo PDF    │    │
│  │  POST /generate-grafs   → Gráficos matplotlib    │    │
│  │  POST /build-structure  → Estructura de slides   │    │
│  │  POST /generate-n-emp   → PDF multi-empresa      │    │
│  │  POST /xlsx-to-pdf      → Excel → PDF            │    │ 
│  │  POST /merge-pdfs       → Unir PDFs              │    │
│  │  GET  /files/read       → Leer archivos          │    │
│  │  GET  /files/list       → Listar archivos        │    │
│  │  POST /files/save       → Guardar archivos       │    │
│  │  GET  /health           → Health check           │    │
│  └──────────────────────────────────────────────────┘    │
│  Volume: /data (n8n-pvc)  → datos compartidos            │
│  Volume: /app (libereoffice-code-pvc) → código fuente    │
│  Volume: /tmp (emptyDir)  → archivos temporales          │
└──────────────────────────────────────────────────────────┘