type: patch

La fase 4 de la adopción lee la política del gate de contrato de PR como lo hace la plantilla canónica: desde `base/` cuando el checkout de la base existe (DysTelefonica/team-skills#1313). Fuera de CI no hay `base/`, y antes el comprobador daba `esta es ilegible` aunque el consumer tuviera la política versionada en el repo. Ahora, sin `base/`, lee la copia del repo; con `base/` presente sigue leyendo `base/`, la copia de confianza. Si no existe ninguna de las dos, sigue fallando cerrado.
