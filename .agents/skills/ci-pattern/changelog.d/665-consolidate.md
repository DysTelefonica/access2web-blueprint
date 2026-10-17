type: minor

Nuevo consolidador de fragmentos (#665 fase C): el PR de release suma los changelog.d/ pendientes, calcula la versión (minor si hay algún minor, si no patch), la escribe en SKILL.md, añade la entrada a CHANGELOG.md y borra los fragmentos. Se ejecuta a mano en el PR de release; no corre en CI.
