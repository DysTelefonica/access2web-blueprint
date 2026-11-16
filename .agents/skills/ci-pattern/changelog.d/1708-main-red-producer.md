type: minor

Nuevo asset `runner-controls/main_red.py` (#1708, parte A): el productor del rojo de `main`. Tras un run de CI en la rama por defecto que termina en rojo, abre una sola issue con la etiqueta `prerelease:main-red` (o reutiliza la abierta, con un comentario), con el run, el SHA y la PR fusionada. Un run en verde cierra esas issues. Todo sin intervención humana y sin `gh`: la API va por `github_rest.py`. Cualquier fallo de lectura o escritura sale con 2 y el comando para relanzarlo. Incluye la plantilla `main-red.template.yml` para el workflow del consumer. Refs #1708.
