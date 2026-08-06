# scripts/setup-rclone.md

# Configurar rclone con Cloudflare R2

Los binarios legacy (`.accdb`, `.codegraph-vba/`, `resources/`) viven en **Cloudflare R2** y se descargan/suben con `rclone`. Esta guía explica el setup único por máquina.

## 1. Instalar rclone

```powershell
# Windows (Chocolatey)
choco install rclone

# O descargar: https://rclone.org/downloads/
```

Verificar:
```powershell
rclone --version
```

## 2. Obtener credenciales R2

En [dash.cloudflare.com](https://dash.cloudflare.com/) → R2 → **Manage R2 API Tokens** → **Create API Token**:
- **Permissions**: Object Read & Write
- **Bucket**: Apply to specific bucket → `access2web-staging-binaries` (o All)
- **TTL**: sin expiry (recomendado para uso continuo)

Click **Create API Token**. La pantalla siguiente muestra **una sola vez**:
- **Access Key ID** — ej: `a1b2c3d4e5f6...`
- **Secret Access Key** — ej: `g7h8i9j0k1l2...`
- **Endpoint** — ej: `https://aa3512b05860af04e1adf4bcd990fe94.r2.cloudflarestorage.com`

Los tokens viejos **siguen funcionando** hasta que los revoqués. Si necesitás rotar, generá uno nuevo, actualizá las configs, y revocá el viejo.

## 3. Configurar el remote en rclone

### Opción A — Comando (recomendado)

```powershell
rclone config create cloudflare-r2 s3 `
  provider Cloudflare `
  access_key_id "<ACCESS_KEY_ID>" `
  secret_access_key "<SECRET_ACCESS_KEY>" `
  endpoint "https://<ACCOUNT_ID>.r2.cloudflarestorage.com" `
  region auto
```

Reemplazá `<ACCESS_KEY_ID>`, `<SECRET_ACCESS_KEY>` y `<ACCOUNT_ID>` con los valores del paso 2.

### Opción B — Manual

Editá `C:\Users\adm1\.config\rclone\rclone.conf` (Windows) o `~/.config/rclone/rclone.conf` (Linux/Mac):

```ini
[cloudflare-r2]
type = s3
provider = Cloudflare
access_key_id = <ACCESS_KEY_ID>
secret_access_key = <SECRET_ACCESS_KEY>
endpoint = https://<ACCOUNT_ID>.r2.cloudflarestorage.com
region = auto
```

## 4. Verificar acceso

```powershell
rclone listremotes                        # tiene que aparecer "cloudflare-r2"
rclone lsd cloudflare-r2:                 # lista los buckets
rclone ls cloudflare-r2:access2web-staging-binaries/ --max-depth 1   # contenido
```

Si ves los buckets, todo OK.

## 5. Primera descarga (en una máquina nueva)

```powershell
pwsh -File scripts/setup-staging.ps1
```

Esto descarga los binarios de las 8 apps (~870 MB) a `data/staging/<app>/{frontend,backend,resources,.codegraph-vba}`.

## 6. Subir cambios (si modificás binarios)

```powershell
pwsh -File scripts/sync-to-r2.ps1 -App condor        # solo Condor
pwsh -File scripts/sync-to-r2.ps1                   # todas las 8 apps
```

Por default, rclone **skip los archivos que ya están iguales** (compara modtime + size). Si querés forzar re-upload, agregá `--no-check-dest` al `rclone copy` dentro del script.

## Troubleshooting

- **"Remote cloudflare-r2 not configured"** → Volvé al paso 3.
- **"No se puede acceder al bucket"** → Verificá el endpoint (debe ser `https://<ACCOUNT_ID>.r2.cloudflarestorage.com`, no `https://r2.cloudflarestorage.com`).
- **"403 Forbidden"** → Las credenciales están mal o el token no tiene permisos. Rotá en Cloudflare → R2 → Manage R2 API Tokens.
- **Push a GitHub falla con timeout** → No es un problema de rclone; ver `.gitignore` para confirmar que los binarios están excluidos.
