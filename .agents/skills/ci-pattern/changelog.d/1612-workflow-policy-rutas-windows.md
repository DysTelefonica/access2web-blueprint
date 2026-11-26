type: patch

El gate de política de workflows ya reconoce en Windows un workflow reutilizable llamado desde un workflow de `pull_request` (#1270). `workflow_id` solo cortaba la ruta por `.github/workflows/` con `/`; con una ruta de Windows (`.github\workflows\callee.yml`) el llamado no casaba con el `uses:` de quien lo llama, no heredaba el evento `pull_request` y un job que leía `secrets.TEAM_SKILLS_READ_TOKEN` pasaba sin hallazgo: un falso negativo del gate, no un defecto del test. Ahora la ruta se normaliza a `/` antes de compararla.

El test `test_workflow_policy.py` añade un caso con rutas de Windows que falla igual en Linux sin el arreglo. Además, el bloque `if __name__ == "__main__"` estaba a mitad de fichero, así que al lanzarlo como script se saltaban las clases posteriores; ahora va al final y corren todas.
