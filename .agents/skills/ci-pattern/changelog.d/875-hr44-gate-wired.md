type: patch

La línea HR-44 («fragmento válido en `changelog.d/`») de la sección «Pull requests» del doc operativo sólo se emite cuando algún `.github/workflows/*.yml` del consumer cablea `check_changelog_fragment.py`. Sin el cableado la línea desaparece, para no obligar a un paso que nadie verifica. La detección lee los workflows: no hay campo nuevo en el host-contract ni en `parameters.schema.json` (#875, oleada 2 de #873).
