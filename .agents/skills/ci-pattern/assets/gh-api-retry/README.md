# gh-api-retry

Reintenta una llamada de la CLI `gh` que **falla**, sin tocar el fail-closed
del gate que la usa (#1210, #1214).

## La regla (una sola, reutilizable)

- Se reintenta **sólo un fallo de la llamada**: `gh` sale distinto de 0.
- Una **respuesta válida** se devuelve tal cual y **nunca** se reintenta. Una
  lista vacía puede significar «sin PR» o «cero issues abiertas»: reintentarla
  debilitaría el gate, no lo haría más robusto.
- Agotados los intentos, sale con `1` y el último error por `stderr`.

## Uso

```bash
response="$(assets/gh-api-retry/gh-api-retry.sh api --paginate "repos/$repo/commits/$sha/pulls")" \
  || { echo "ERROR: la consulta de PR para $sha falló; se rechaza el push (fail closed)." >&2; exit 2; }
```

Funciona con cualquier subcomando de `gh` (`api`, `issue list`, `pr list`…):
los argumentos se pasan tal cual.

## Entorno

| Variable | Por defecto | Para qué |
|---|---|---|
| `GH_API_RETRY_INTENTOS` | 3 | intentos totales; `1` desactiva el reintento |
| `GH_API_RETRY_ESPERA_S` | 1 | segundos de espera entre intentos |

## Streams

La salida de `gh` va a `stdout` y su `stderr` a `stderr`, sin mezclarse: un
aviso de `gh` no puede corromper el JSON que el gate va a parsear.

## Qué NO hace

No reintenta una respuesta válida, no cambia el fail-closed del gate que la
llama y no reintenta nada que no sea una llamada a la API.
