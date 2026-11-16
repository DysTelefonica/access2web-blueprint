# prerelease-deploy — comprobación diaria de deploy (#1627, slice 7)

`check_daily_deploy.py` exige al menos un deploy exitoso a producción por día UTC. Si no
lo hay, abre una issue de máxima prioridad con la etiqueta fija `prerelease:night-red`.

## Uso

```bash
python3 check_daily_deploy.py --repo owner/name                  # hoy (UTC), entorno production
python3 check_daily_deploy.py --repo owner/name --date 2026-11-13
python3 check_daily_deploy.py --repo owner/name --environment staging --dry-run
```

## Qué cuenta como deploy exitoso

Un GitHub Deployment del entorno (`--environment`, por defecto `production`) creado ese
día UTC cuyo estado más reciente es `success`. Con un deployment del día que tenga como
último estado `failure`, aunque tuviera un `success` anterior, la comprobación falla.

## Códigos de salida

| Código | Significado |
|---|---|
| 0 | Hubo deploy exitoso ese día |
| 1 | No lo hubo: se abre la issue (o ya hay una abierta, o es ensayo) |
| 2 | La API no se pudo leer o la fecha es inválida: fallo cerrado, sin issue |

Si ya hay una issue abierta con la misma etiqueta y el mismo título, no se abre otra.
La etiqueta `prerelease:night-red` debe existir en el repositorio; el script no la crea.

## Pruebas

`python3 -m unittest discover -s tests -p 'test_check_daily_deploy*.py'`.

Las pruebas apuntan `GITHUB_API_URL` al servidor stub local
(`runner-controls/tests/_github_stub.py`) y no tocan la red. El asset habla con la
API por `runner-controls/github_rest.py`, no con el CLI `gh`: el runner del consumer
no lo trae (#1691). El token sale de `GITHUB_TOKEN` (o `GH_TOKEN`) del entorno.
