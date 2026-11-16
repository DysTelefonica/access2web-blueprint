# gh-api-retry

La regla del reintento de las llamadas a la API, en un solo sitio (#1210, #1220, #1592).

## La regla

Se reintenta lo transitorio con **backoff exponencial**: 4 intentos, esperas
de 2, 4 y 8 s. Un 4xx distinto de 429 no se reintenta. Agotados los intentos,
se falla cerrado (HR-3).

Un 4xx recibiría la misma respuesta otra vez; el 429 (límite de peticiones)
pasa con el tiempo. El 403 del límite secundario no se reintenta: GitHub lo da
con el mismo código que un permiso que falta.

Una respuesta válida **nunca** se reintenta: una lista vacía puede significar «sin PR»
o «cero issues abiertas», y reintentarla debilitaría el gate en vez de
protegerlo.

## La tabla es el contrato

`casos.tsv` contiene esas filas, con sus esperas (los tests las observan sin
dormir). **Las dos implementaciones la recorren en
sus tests**:

- este asset, en bash, que envuelve la CLI `gh` (`gh-api-retry.sh`);
- el cliente Python de la canon, `runner-controls/github_rest.py`, que usan los
  gates de `pr-contract`, `release-gate` y `runner-controls`.

Si una implementación se desvía de la tabla, su suite falla. No hace falta
comparar las dos a mano: la tabla es la expectativa de las dos.

## Cómo se usa

```bash
# Bash: envuelve cualquier subcomando de gh
response="$(assets/gh-api-retry/gh-api-retry.sh api --paginate "repos/$repo/commits/$sha/pulls")" \
  || { echo "ERROR: la consulta de PR para $sha falló; se rechaza el push (fail closed)." >&2; exit 2; }
```

```python
# Python: el cliente ya reintenta solo
data = github_rest.get_json(f"repos/{repo}/commits/{sha}/pulls")
```

## Entorno

| Variable | Por defecto | Para qué |
|---|---|---|
| `CI_PATTERN_GH` | el `gh` del PATH | el `gh` a ejecutar, como en el resto de assets (#1609, #1632); una ruta `.py` corre con `python3` |
| `GH_API_RETRY_INTENTOS` | 4 | intentos totales; `1` desactiva el reintento |
| `GH_API_RETRY_ESPERA_S` | 2 | primera espera en segundos; cada reintento la dobla; `0` = sin esperas (tests) |

## Detalles que no son obvios

- **Cómo se detecta un JSON malformado en bash**: sin parser y sin dependencias,
  la API responde objetos o arrays, así que un cuerpo vacío o que no empieza por
  `{` o `[` se trata como malformado. Es un criterio de forma, no de parseo, y
  por eso este helper es para consultas JSON de la API.
- **Escrituras: solo las idempotentes** (#1592). El cliente Python reintenta
  las lecturas (`get_json`) y, con `call_con_reintento`, las escrituras que
  repetidas dejan el mismo estado: el `PUT .../merge` con `sha` del merge
  gobernado (si el primero llegó a fusionar, el reintento recibe un 4xx y falla
  cerrado; nunca fusiona otra cabeza), el `PATCH` que cierra una issue y la
  consulta GraphQL (un `POST` que solo lee). Las escrituras **no idempotentes**
  —el `POST` de un comentario, de una issue de incidente o de un
  `workflow_dispatch`— van por `call` y hacen **un solo intento**: repetidas
  tras un 5xx que sí se aplicó, duplicarían el comentario, la issue o el run, y
  un duplicado es peor que un rojo que falla cerrado. Por eso tampoco hay
  comprobación previa: añadiría una lectura más que también puede fallar. El
  helper bash no se usa para escrituras. Un fallo de configuración (sin token)
  no se reintenta.
- **Cómo sabe el helper bash que es un 4xx**: `gh api` escribe
  `gh: <mensaje> (HTTP nnn)` en stderr; sin código (error de red), se reintenta.
- **Quién más usa la regla**: el gate `issue-spec` (`check_issue_spec.py`, el
  job `pr-gate`) lee el cuerpo del PR y las issues con la regla del cliente
  compartido (`github_rest.con_reintento`); antes llamaba a `urlopen` con un
  único intento (#1592, APAP_WEB job 131271799279). `pr-contract` y
  `required-jobs` leen con `get_json`, que ya reintenta.
- **Streams separados**: la salida de `gh` va a `stdout` y su `stderr` a
  `stderr`, sin mezclarse: un aviso de `gh` no puede corromper el JSON que el
  gate va a parsear.

## Qué NO hace

No reintenta una respuesta válida, no cambia el fail-closed del gate que llama
y no reintenta nada que no sea una llamada a la API.
