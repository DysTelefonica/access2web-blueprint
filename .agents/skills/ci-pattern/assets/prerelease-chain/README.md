# prerelease-chain — apertura de eslabones de cadena (#1627, slice 7)

`open_chain_link.py` abre un eslabón de una cadena de PRs con la forma de la regla 5
del perfil pre-release.

## Uso

```bash
python3 open_chain_link.py --repo owner/name --issue 1627 --head feat/x-link-2 \
  --title "eslabon 2" --body-file cuerpo.md          # eslabón intermedio
python3 open_chain_link.py ... --tip                  # la punta: Closes #N
python3 open_chain_link.py ... --dry-run              # imprime el comando, no crea nada
```

## Reglas que aplica

- Intermedio: etiqueta `chain:partial`, primera línea `Refs #N`, sin `Closes`, `Fixes` ni `Resolves` en el cuerpo.
- Punta (`--tip`): primera línea `Closes #N`, sin etiqueta de cadena.
- Como máximo 3 eslabones abiertos con `chain:partial`; un cuarto se rechaza (exit 1).

## Códigos de salida

| Código | Significado |
|---|---|
| 0 | Abierto, o ensayo correcto |
| 1 | Rechazado por la política (ya hay 3 eslabones abiertos) |
| 2 | Entrada inválida o gh ilegible: fallo cerrado |

## Pruebas

`python3 tests/test_open_chain_link.py`. gh se sustituye por `tests/fake_gh.py` mediante
`PRERELEASE_GH_CMD` (JSON); las pruebas no tocan la red.
