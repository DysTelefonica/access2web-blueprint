type: minor

#772: la plantilla de PR del canon trae de serie los cinco enlaces `docs/codebase/maintainer-playbook.md#feature|bug|refactor|docs|chore` bajo «Documentación» (HR-64), y la plantilla de guía de mantenimiento añade el encabezado `### Chore`. Un test fija que cada enlace de la plantilla resuelve a un encabezado de la guía. Los consumers reciben la plantilla por repropagación del catálogo (su hash en `.governance-manifest.json`, HR-46).
