# Gate Liveness — Aviso de gate fail-closed bloqueado (F3, #396, HR-3/HR-21)

## Problema

Un gate fail-closed que se queda bloqueado (healthz down, e2e fallido, deploy
tumbado) no produce ninguna seÃ±al externa. El CI no avisa que el gate lleva
horas/dÃ­as parado. Ejemplo: el VPS tumbado el 2026-10-05 dejÃ³ 4 dÃ­as de
producciÃ³n caÃ­da y 74 commits mergeados sin desplegar (#1290).

## Regla

**HR-21**: un gate fail-closed bloqueado `MUST` producir una seÃ±al externa
(issue, correo o webhook) con un umbral de tiempo declarado. Desbloquear
exige intervenciÃ³n humana; enterarse no.

## Uso

```bash
python3 assets/gate-liveness/check_gate_liveness.py --repo <repo> [--threshold <horas>]
```

- `--threshold`: umbral de horas antes de alertar (default: 24)
- Exit 0: gates fail-closed detectados, liveness configurado
- Exit 1: gates fail-closed sin liveness configurado (FINDING)
- Exit 2: no se detectaron gates fail-closed (SKIP)

## ImplementaciÃ³n

El asset escanea los workflows YAML en `.github/workflows/` y detecta jobs
cuyos nombres contienen palabras clave (`gate`, `security`, `e2e`, `deploy`,
`health`, `smoke`, `verify`, etc.) o que estÃ¡n en `needs` de un job de
despliegue. Luego busca si el consumer tiene configuraciÃ³n de liveness:

- Directorio `assets/gate-liveness/` existente, o
- MenciÃ³n de `liveness` en un workflow, o
- Archivo JSON con key `liveness_threshold` / `alert_after` / `stale_threshold`

## Configurar

Para resolver el FINDING, crea un directorio `assets/gate-liveness/` en el
repo con un README o configuraciÃ³n de umbral, o agrega un job `gate-liveness`
al workflow que verifique periÃ³dicamente que los gates estÃ©n vivos.
