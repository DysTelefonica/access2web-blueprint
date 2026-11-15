# Runbook: thin pool de Proxmox

Alerta: `[<host>] ALERTA thin pool <vg>/<lv> al <N>%`, enviada por `thinpool-guard` en `pve`, `pve2` o `pve3`.

## Por qué importa

Los discos de los CT y las VM suman más que el thin pool que los aloja: el pool está sobreaprovisionado. Si llega al 100 %, todas las escrituras del host fallan a la vez y caen todos los CT, no solo el que más ocupa.

## 1. Medir

```sh
lvs -o lv_name,lv_size,data_percent,metadata_percent,origin,pool_lv <vg>
```

- La fila del pool (`data`) da el uso real. Las filas `vm-<id>-disk-N` dan cuánto ocupa cada disco; las `snap_vm-<id>-disk-N_<nombre>` son instantáneas.
- Una instantánea no muestra su ocupación propia: la estimación es la fila del pool menos la suma de los discos (`lv_size × data_percent`).

## 2. Localizar

```sh
pct list
for ct in $(pct list | awk 'NR>1 {print $1}'); do echo "CT $ct"; pct listsnapshot "$ct"; done
systemctl list-units --no-legend 'cadete-ephemeral@*'
```

Clasifique cada CT que crece:

| Tipo de CT | Cómo reconocerlo | Qué hacer |
|---|---|---|
| Runner efímero de cadete | Tiene una unidad `cadete-ephemeral@<ct>` activa y una instantánea `clean` | Si está en un job, esperar: al acabar vuelve a `clean` y libera el espacio. Ver el job con `journalctl -u cadete-ephemeral@<ct> -n 30`. |
| CT persistente | Sin unidad efímera | Los bloques que su sistema de ficheros ya liberó no vuelven al pool solos: proponer `pct fstrim <ct>`. |
| Instantánea antigua | `pct listsnapshot` muestra instantáneas distintas de `clean` (por ejemplo `pre-ephemeral`) | Proponer `pct delsnapshot <ct> <nombre>`. Es irreversible: confirmar. |

## 3. Acciones, de menor a mayor impacto

| Acción | Efecto | Reversible | Confirmación |
|---|---|---|---|
| Esperar al final del job de un CT efímero | La vuelta a `clean` libera el disco del job | Sí | No |
| `pct fstrim <ct>` en un CT persistente | Devuelve al pool los bloques libres del CT | Sí, no borra datos | Sí |
| `pct delsnapshot <ct> <nombre>` (nunca `clean`) | Libera lo que solo esa instantánea retiene | No | Sí |
| `systemctl stop cadete-ephemeral@<ct>` | El CT deja de tomar jobs; la CI de cadete pierde ese slot | Sí, con `systemctl start` | Sí |
| `lvextend --poolmetadatasize +<tamaño> <vg>/<lv>` | Amplía los metadatos del pool | No | Sí, y por separado |

Una medición de referencia: borrar las dos `pre-ephemeral` de los CT 312 y 313 en `pve3` liberó unos 11,7 GiB (del 72,8 % al 64,5 %).

## 4. Verificar

Repita `lvs` sobre el pool. Bajo el umbral, `thinpool-guard` envía `[<host>] OK thin pool ...` en su siguiente pasada. El log de sus pasadas está en `/var/log/thinpool-guard.log`.
