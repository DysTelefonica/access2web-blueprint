# ci-red: el rojo programado tiene dueño (HR-89)

Un run programado (`schedule`) en rojo sobre la rama por defecto abre una issue
de seguimiento con la etiqueta `ci-red`, asignada al coordinador que declara el
contrato del host. Si esa issue ya está abierta, el nuevo rojo la comenta. El
primer run programado en verde la comenta y la cierra. Hay una issue por
workflow, con el título `[ci-red] <workflow>`.

`report_ci_red.py` se copia en el consumer y solo usa la biblioteca estándar y
`gh`. No invoca la CLI del patrón (lección #342).

## Requisitos del contrato del host

`.github/host-contract.json` declara las dos cosas siguientes. Si falta
alguna, el script sale con 2 y no llama a `gh`:

```json
{
  "coordinator": "<login del coordinador>",
  "labels": [
    {"name": "ci-red", "class": "host-enforced",
     "description": "Run programado en rojo sobre la rama por defecto; la abre y la cierra el aviso (HR-89)",
     "color": "b60205"}
  ]
}
```

Cree la etiqueta en el host con `scripts/sync-governance-labels.sh`: el script
emite el comando `gh label create` y una persona con permisos lo ejecuta.

## Pasos del workflow programado

Ponga los dos pasos al final del job. El job necesita `issues: write`. El
`if:` no comprueba la rama: de eso se encarga el script, que con otro evento o
en otra rama responde «nada que hacer» y no llama a `gh`.

```yaml
    permissions:
      contents: read
      issues: write
    steps:
      # ... pasos del job ...
      - name: ci-red (rojo programado, HR-89)
        if: failure() && github.event_name == 'schedule'
        env:
          GH_TOKEN: ${{ github.token }}
        run: >-
          python3 .agents/skills/ci-pattern/assets/ci-red/report_ci_red.py
          --outcome failure
          --default-branch "${{ github.event.repository.default_branch }}"
          --host-contract .github/host-contract.json
      - name: ci-red (verde programado, HR-89)
        if: success() && github.event_name == 'schedule'
        env:
          GH_TOKEN: ${{ github.token }}
        run: >-
          python3 .agents/skills/ci-pattern/assets/ci-red/report_ci_red.py
          --outcome success
          --default-branch "${{ github.event.repository.default_branch }}"
          --host-contract .github/host-contract.json
```

Con varios jobs en el workflow, los pasos van en un job final. Ese job lleva
`needs` con todos los demás y `if: always() && github.event_name == 'schedule'`.
En cada paso, el `if:` lee el resultado de los jobs previos:
`contains(needs.*.result, 'failure')` para el rojo y
`!contains(needs.*.result, 'failure') && !contains(needs.*.result, 'cancelled')`
para el verde. Así cada run deja un solo aviso por workflow.

## Códigos de salida

| Código | Significado |
| --- | --- |
| `0` | Aviso dejado, issue cerrada o nada que hacer |
| `1` | `gh` no pudo dejar el aviso (HR-3): el paso sale en rojo, nunca con un 0 silencioso |
| `2` | Uso, identidad del run (`GITHUB_REPOSITORY`, `GITHUB_WORKFLOW`) o contrato del host inservibles |

## Tests

`tests/test_report_ci_red.py` ejecuta el script contra un `gh` simulado y nunca
llama al real. En Windows usa un lanzador `gh.cmd`; en POSIX, un `gh`
ejecutable. Un test de guarda falla si `gh` no resuelve al simulado.
