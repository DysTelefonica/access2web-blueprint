# gh-api-retry

La regla del reintento de las consultas a la API, en un solo sitio (#1210, #1220).

## La regla (cuatro casos)

| Caso | Qué hace la API | Reintenta | Intentos | Resultado |
|---|---|---|---|---|
| fallo de llamada | error de red o HTTP | sí | 3 | la llamada responde bien |
| JSON malformado | el cuerpo no es JSON | sí | 3 | la llamada responde bien |
| respuesta válida vacía | `[]` o cuerpo vacío válido | **no** | 1 | se devuelve tal cual |
| intentos agotados | siempre falla | sí | 3 | **falla cerrado** |

Una respuesta válida **nunca** se reintenta: una lista vacía puede significar «sin PR»
o «cero issues abiertas», y reintentarla debilitaría el gate en vez de
protegerlo.

## La tabla es el contrato

`casos.tsv` contiene esas cuatro filas. **Las dos implementaciones la recorren en
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
| `GH_API_RETRY_INTENTOS` | 3 | intentos totales; `1` desactiva el reintento |
| `GH_API_RETRY_ESPERA_S` | 1 | segundos de espera entre intentos |

## Detalles que no son obvios

- **Cómo se detecta un JSON malformado en bash**: sin parser y sin dependencias,
  la API responde objetos o arrays, así que un cuerpo vacío o que no empieza por
  `{` o `[` se trata como malformado. Es un criterio de forma, no de parseo, y
  por eso este helper es para consultas JSON de la API.
- **El cliente Python reintenta lecturas, no escrituras**: un POST repetido
  podría duplicar el efecto (por ejemplo, abrir dos incidencias). El helper
  bash tampoco se usa para escrituras.
- **Streams separados**: la salida de `gh` va a `stdout` y su `stderr` a
  `stderr`, sin mezclarse: un aviso de `gh` no puede corromper el JSON que el
  gate va a parsear.

## Qué NO hace

No reintenta una respuesta válida, no cambia el fail-closed del gate que llama
y no reintenta nada que no sea una llamada a la API.
