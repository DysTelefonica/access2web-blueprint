---
name: infra-disk-alerts
description: "Trigger: alerta de disco, aviso del bot de Telegram, ALERTA thin pool, thin pool lleno, disk at, disk guard ran, vps-disk-guard, thinpool-guard, disco lleno en rosa u oracle, reinstalar thinpool-guard. Atiende las alertas de disco de los hosts de infraestructura: reconoce el mensaje, diagnostica en solo lectura en el host que lo envía y propone la acción, que se confirma antes de escribir."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-10-30
  scope: ['universal', 'infra']
  auto_invoke: ['handling an infrastructure disk alert']
  tiers: ['universal', 'infra']
---

# infra-disk-alerts

Los hosts de infraestructura avisan por Telegram, a través del bot `@ardelperal_backup_bot`, cuando se les llena el disco. Esta skill dice qué hacer cuando el usuario pega uno de esos mensajes: identificar quién lo envía, averiguar qué llena el disco y proponer cómo liberarlo sin romper los runners ni los servicios.

## §1 Activation

Cargue esta skill cuando:

- El usuario pegue un mensaje del bot con la forma `[<host>] ALERTA thin pool <vg>/<lv> al <N>%` o `<host>: disk at <N>% (<M>MB free) after the disk guard ran.`
- El usuario pregunte por un thin pool lleno en `pve`, `pve2` o `pve3`, o por el disco raíz de `rosa` u `oracle`.
- Haya que reinstalar `thinpool-guard` en un host Proxmox o cambiar su umbral.

**Do NOT load when:**

- La tarea sea provisionar o migrar runners en el VPS de Oracle: eso es `oracle-vps-github-runners`.
- Haya que cambiar la política de limpieza de `vps-disk-guard.sh`: la gestiona Ansible desde el repositorio `vps-oracle`.
- El mensaje sea `[<host>] OK thin pool ...`: es la recuperación y no requiere acción; confírmela con una medición y termine.

## §2 Hard Rules

- **HR-1** — Identifique el origen antes de actuar. El mensaje determina el host y el runbook con la tabla de §3; un mensaje que no encaje en ninguna fila MUST tratarse como `alert_source: "unknown"` y MUST NOT disparar ninguna escritura.
- **HR-2** — Conéctese solo al host que nombra la alerta y solo por su alias SSH (`pve`, `pve2`, `pve3`, `oracle`, `rosa`). En Windows use `C:\Windows\System32\OpenSSH\ssh.exe -o BatchMode=yes <alias>`; el `ssh` de Git Bash falla la verificación de la clave del host. Pegar la alerta pidiendo que se atienda autoriza el diagnóstico de solo lectura en ese host; MUST NOT extenderse a otros hosts sin autorización.
- **HR-3** — Diagnostique en solo lectura: `lvs`, `pct list`, `pct listsnapshot`, `df`, `du`, `docker system df`, `journalctl`, `tail` de logs. Toda escritura (borrar instantáneas, `pct fstrim`, parar unidades, `docker rm`, `docker image rm`, ampliar volúmenes) MUST proponerse con su efecto y su reversibilidad, y MUST esperar la confirmación del usuario.
- **HR-4** — No borre nunca la instantánea `clean` de un CT con unidad `cadete-ephemeral@<ct>`: el controlador vuelve a ella después de cada job y, sin ella, el runner deja de funcionar.
- **HR-5** — No ejecute `docker volume prune`, `docker system prune --volumes` ni `docker system prune -a` en `rosa` u `oracle`: los volúmenes guardan los datos de servicios de Coolify y las imágenes con tag son el destino de un rollback.
- **HR-6** — No añada crons de limpieza ni edite `vps-disk-guard.sh` en el host: su dueño es Ansible (`vps-oracle`, `playbooks/disk-guard.yml`). Un cambio de política se hace allí.
- **HR-7** — No imprima ni copie al repositorio el token del bot ni la contraseña de aplicación de Gmail. Viven en `/root/.vps-disk-guard-telegram.env` y `/root/.gmail-app-password` (0600, root). Para comprobarlos, mida longitud y formato; para copiarlos entre hosts, use una tubería SSH a SSH sin mostrarlos.
- **HR-8** — Cierre con una medición nueva, no con la suposición de que la acción liberó espacio. Si el uso sigue por encima del umbral, el estado es `blocked` y se informa de qué lo ocupa.

## §3 Decision Gates

| Condition | Action |
|---|---|
| Mensaje `[<host>] ALERTA thin pool <vg>/<lv> al <N>%` | Runbook de thin pool: `references/runbook-thin-pool.md`, en el host `<host>` (`pve`, `pve2` o `pve3`). |
| Mensaje `[<host>] OK thin pool ...` | Recuperación. Medir con `lvs` para confirmarla; ninguna acción más. |
| Mensaje `[<host>] thinpool-guard: prueba` | Prueba del canal. Ninguna acción. |
| Mensaje `vnic-rosa: disk at ...` | Runbook de disco raíz en el alias `rosa`: `references/runbook-root-disk.md`. |
| Mensaje `vinc-ardelperal: disk at ...` | Runbook de disco raíz en el alias `oracle`: `references/runbook-root-disk.md`. |
| El thin pool crece por un CT efímero con un job en curso | Esperar al final del job y volver a medir: la vuelta a `clean` libera su espacio. |
| El thin pool supera el 95 % con jobs en curso | Proponer parar `cadete-ephemeral@<ct>` del CT que más crece, para que no tome otro job. Reduce la capacidad de CI: confirmar. |
| Los metadatos del pool superan el umbral | Proponer `lvextend --poolmetadatasize`; es irreversible. Confirmar y no combinarlo con otra acción. |
| El disco raíz sigue alto después del guard | Lo que queda es lo que el guard no debe borrar: identificar el consumidor con `du` y proponer la acción concreta. |
| Un host no responde por SSH | `status: "blocked"`. Informar del host y no probar otros caminos. |
| `thinpool-guard` falta en un host Proxmox o hay que cambiar su umbral | Reinstalar con `references/thinpool-guard-install.md`. |

## §4 Execution Steps

1. **Identificar.** Clasifique el mensaje con §3 y extraiga host, recurso y porcentaje. Traduzca el nombre del host a su alias SSH: `pve`, `pve2`, `pve3`, `vnic-rosa` → `rosa`, `vinc-ardelperal` → `oracle`.
2. **Medir.** Conéctese al alias (HR-2) y repita la medición: `lvs` para el thin pool, `df -h /` para el disco raíz. El valor puede haber cambiado desde el aviso.
3. **Localizar.** Siga el runbook que indica §3 hasta encontrar qué ocupa el espacio: qué LV o instantánea en el pool, qué directorio en el disco raíz.
4. **Proponer.** Presente la acción concreta con su efecto, cuánto liberaría y si es reversible (HR-3). Espere la confirmación.
5. **Aplicar.** Ejecute solo lo confirmado.
6. **Verificar.** Vuelva a medir (HR-8). Si baja del umbral, `thinpool-guard` envía el aviso de recuperación en la siguiente pasada (cada 10 minutos).
7. **Informar.** Devuelva el Output Contract.

## §5 Output Contract

| Key | Type | Description |
|---|---|---|
| `status` | `"resolved" \| "waiting" \| "blocked" \| "no_action"` | `resolved`: medición final bajo el umbral; `waiting`: se espera al final de un job; `blocked`: sigue alto o el host no responde; `no_action`: recuperación o prueba. |
| `alert_source` | `"thinpool-guard" \| "vps-disk-guard" \| "unknown"` | Quién envió el mensaje. |
| `host` | string | Alias SSH del host (`pve`, `pve2`, `pve3`, `rosa`, `oracle`). |
| `resource` | string | `<vg>/<lv>` del thin pool o `/` para el disco raíz. |
| `usage_pct_before` | number | Uso medido al empezar, no el del mensaje. |
| `usage_pct_after` | number \| null | Uso medido al terminar; `null` si no se aplicó nada. |
| `cause` | string | Qué ocupa el espacio (LV, instantánea, directorio, imagen). |
| `actions_taken` | string[] | Escrituras ejecutadas tras confirmación; vacío si ninguna. |
| `actions_proposed` | string[] | Acciones pendientes de confirmación. |

## §6 Anti-patterns

| Symptom | Fix |
|---|---|
| Se borra la instantánea `clean` de un CT efímero para liberar espacio | Prohibido (HR-4). Borrar solo otras instantáneas, como `pre-ephemeral`, y con confirmación. |
| Se lanza `docker system prune -a --volumes` en `rosa` | Prohibido (HR-5). Identificar con `docker system df -v` qué ocupa y proponer el borrado concreto. |
| Se da por resuelto sin volver a medir | Medir con `lvs` o `df` después de la acción (HR-8). |
| Se hace `cat` del fichero del token para comprobarlo | Comprobar longitud y formato sin mostrarlo (HR-7). |
| Se añade un cron de limpieza en `oracle` | Cambiar la política en `vps-oracle` (HR-6). |
| Se conecta a todos los hosts para buscar el problema | Solo el host que nombra la alerta (HR-2). |

## §7 Companion skills

| Skill | Cargar junto cuando |
|---|---|
| `oracle-vps-github-runners` | La alerta de disco raíz apunta a los runners de `oracle` o `rosa` (`_work`, Docker interno de un runner). |


El ciclo de los runners efímeros de cadete (instantánea `clean`, unidad `cadete-ephemeral@<ct>`, reparto de jobs) está documentado en el repositorio `DysTelefonica/cadete`, `docs/infrastructure/runners-efimeros.md`. Léalo antes de parar o reanudar una de esas unidades.
