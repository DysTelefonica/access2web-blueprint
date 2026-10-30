type: patch

El contrato de PR (`assets/pr-contract/check_pr_contract.py`) admite rutas relativas que empiezan por punto en la sección «Tests que prueban el cierre» (DysTelefonica/team-skills#1355). El parser de HR-53 descartaba antes una entrada válida como `.agents/skills/<skill>/tests/<caso>.py`, de modo que el gate emitía un aviso aunque la ruta estuviera versionada. Ahora la entrada se reconoce y se valida con las reglas existentes.
