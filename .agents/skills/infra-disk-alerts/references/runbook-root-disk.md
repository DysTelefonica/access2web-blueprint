# Runbook: disco raíz de rosa u oracle

Alerta: `<hostname>: disk at <N>% (<M>MB free) after the disk guard ran. Something it must not delete is filling /.`, enviada por `vps-disk-guard.sh`.

| Hostname en el mensaje | Alias SSH | Usuario |
|---|---|---|
| `vnic-rosa` | `rosa` | sin root: anteponga `sudo -n` |
| `vinc-ardelperal` | `oracle` | sin root: anteponga `sudo -n` |

## Qué significa

El guard corre cada 10 minutos y ya ha borrado todo lo que puede borrar sin riesgo: `/tmp` sin uso, el `_work` de los runners inactivos, la caché de build y las imágenes sin tag de Docker, los logs grandes de contenedores, la caché de apt y el journal antiguo. Avisa solo si, después de todo eso, el disco sigue por encima del 88 %. Lo que llena el disco es, por tanto, algo que el guard tiene prohibido tocar.

## 1. Medir y leer al guard

```sh
df -h /
sudo -n tail -40 /var/log/vps-disk-guard.log
```

El log dice qué recuperó la última pasada y si algún paso falló.

## 2. Localizar

```sh
sudo -n du -xh --max-depth=1 / 2>/dev/null | sort -h | tail -8
sudo -n docker system df -v | head -40
sudo -n du -sh /data/runners/* 2>/dev/null | sort -h | tail
```

| Lo que ocupa | Por qué el guard no lo borra | Qué proponer |
|---|---|---|
| Volúmenes de Docker | Guardan los datos de servicios de Coolify | Identificar el servicio dueño del volumen. Borrar solo un volumen huérfano confirmado, nunca un `prune`. |
| Imágenes con tag | Son destino de rollback | Borrar con `docker image rm <imagen>:<tag>` las versiones antiguas que nadie despliega, confirmando cada una. |
| `_work` de un runner en un job | El runner está trabajando | Esperar al final del job. |
| Logs fuera de Docker (`/var/log`) | Fuera de su política | Ver el fichero concreto y proponer rotarlo o truncarlo. |
| Datos de una aplicación | No es caché | Escalar al usuario: hace falta ampliar el disco o mover los datos. |

Para lo que es propio de los runners (el Docker que corre dentro de un runner, `_work`, runners parados), siga `oracle-vps-github-runners`.

## 3. Verificar

`df -h /` otra vez. No hay mensaje de recuperación: `vps-disk-guard.sh` solo avisa al pasar el umbral y repite cada 6 horas mientras siga alto.
