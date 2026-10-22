type: patch

`pr_preflight.py` deja de fallar con `TypeError: 'function' object is not subscriptable` cuando se pasa `--token` con `--repository`: `main()` usaba `fetchers[0]` sobre el fetcher que devuelve `_make_fetcher`, que es una función o `None`. La comprobación remota de `status:approved` de la issue enlazada ya corre con token. Prueba nueva por la CLI con un stub de la API en loopback, con y sin token (#924).
