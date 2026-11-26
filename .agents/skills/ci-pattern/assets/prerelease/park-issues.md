# park_issues.py: aparcar y desaparcar issues (#1627, slice 5)

Durante una noche en rojo, el coordinador aparca las issues normales, crea las issues
de arreglo y las hace implementar. Cuando el deploy nocturno sale en verde, desaparca
todo. Este script hace las dos partes deterministas.

## Etiquetas

| Constante | Valor por defecto | Flag para cambiarla | Para qué |
|---|---|---|---|
| `FIX_LABEL` | `prerelease:night-red` | `--fix-label` | issues de arreglo de la noche roja; no se aparcan |
| `PARK_LABEL` | `prerelease:parked` | `--park-label` | marca de las issues aparcadas |

Los nombres están decididos en #1627 y deben existir en el repo antes de usarse. Las issues
de arreglo las crea el coordinador, no este script.

## Uso

```bash
python ci-pattern/assets/prerelease/park_issues.py park --repo OWNER/REPO --night-run RUN_ID --dry-run
python ci-pattern/assets/prerelease/park_issues.py park --repo OWNER/REPO --night-run RUN_ID
python ci-pattern/assets/prerelease/park_issues.py unpark --repo OWNER/REPO --night-run RUN_ID

```

Habla con la API de GitHub por `runner-controls/github_rest.py` (urllib, stdlib), no con el
CLI `gh`: el runner del consumer no lo tiene (#1691). Necesita `GITHUB_TOKEN` (o `GH_TOKEN`)
en el entorno; `GITHUB_API_URL` solo se usa para los tests, contra un servidor local. Sin
token, o sin `runner-controls/github_rest.py` junto al asset, el script no arranca: sale
con 2 y no hace nada.

## Reglas

- `park` solo actúa si la ejecución nocturna terminó `failure`, `cancelled` o
  `timed_out`. Si está en curso, terminó en `success`, o no se puede leer, sale con 2 y
  no escribe nada.
- Ambos exigen que la ejecución sea del workflow nocturno (`--workflow`, por defecto
  `nightly`, sin extensión). Una ejecución de otro workflow, como un `pr-size` rojo, sale
  con 2 y no escribe nada.
- `unpark` solo actúa si la ejecución terminó en `success` Y hay un deploy exitoso a
  `--environment` (por defecto `production`) creado hoy en UTC cuyo estado más reciente es
  `success`. Sin ese deploy sale con 2 y no escribe nada.
- Se leen hasta `--list-limit` issues (5000 por defecto). Si hay más, la lista está
  incompleta y sale con 2.
- Ambos son idempotentes: una segunda ejecución no cambia nada si ya se hizo.
- Una edición rechazada no detiene las demás: sale con 1 y nombra las issues que fallaron.
  Volver a lanzar el comando termina el trabajo.

Salida: la última línea dice el resultado. Códigos: 0 hecho, 1 edición rechazada, 2 retenido.
