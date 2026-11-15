# Instalar o reinstalar thinpool-guard

`assets/thinpool-guard.sh` es la fuente. En cada host Proxmox se instala como `/usr/local/sbin/thinpool-guard`, con un cron cada 10 minutos.

## Comportamiento

- Revisa datos y metadatos de cada thin pool LVM del host.
- Por encima de `THRESHOLD` (85 por defecto) avisa y deja un marcador en `/var/lib/thinpool-guard/`; repite el aviso cada `REPEAT_HOURS` (6) mientras siga alto.
- Al bajar del umbral envía un aviso de recuperación y borra el marcador.
- Canal: Telegram primero; si falla o no está configurado, Gmail por SMTP. Si no sale ningún aviso, no deja marcador y lo reintenta en la siguiente pasada.
- `thinpool-guard --test` envía un aviso de prueba por el mismo canal.

## Secretos que necesita en el host

| Fichero | Contenido | Permisos |
|---|---|---|
| `/root/.vps-disk-guard-telegram.env` | `TELEGRAM_BOT_TOKEN=...` y `TELEGRAM_CHAT_ID=...` (el mismo fichero que lee `vps-disk-guard.sh`) | 0600 root |
| `/root/.gmail-app-password` | Contraseña de aplicación de Gmail de la cuenta remitente | 0600 root |

Si otro host ya los tiene, cópielos sin mostrarlos:

```sh
ssh <origen> 'cat /root/.vps-disk-guard-telegram.env' | ssh <destino> 'umask 077; cat > /root/.vps-disk-guard-telegram.env'
```

## Instalación

Desde la raíz de la skill:

```sh
tr -d '\r' < assets/thinpool-guard.sh | ssh <alias> 'install -m 0755 /dev/stdin /usr/local/sbin/thinpool-guard'
ssh <alias> 'echo "*/10 * * * * root /usr/local/sbin/thinpool-guard >> /var/log/thinpool-guard.log 2>&1" > /etc/cron.d/thinpool-guard && chmod 644 /etc/cron.d/thinpool-guard'
ssh <alias> '/usr/local/sbin/thinpool-guard --test && /usr/local/sbin/thinpool-guard'
```

La salida de `--test` dice por qué canal salió (`sent via telegram` o `sent via gmail`). Para otro umbral, añada `THRESHOLD=<N>` delante del comando en la línea del cron.

## Test

`assets/tests/test-thinpool-guard.sh` prueba el script contra stubs de `lvs`, `curl` y `python3`; corre en la suite unitaria del catálogo.
