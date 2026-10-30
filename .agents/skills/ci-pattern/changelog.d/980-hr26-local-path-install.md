type: patch

HR-26 del gate `workflow-policy` admite una línea `pip install` cuyos destinos son todos rutas locales explícitas (`./app[dev]`, `-e ./app`, `.`, con prefijos `./`, `../`, `/` o `~/`): sus pines viven en el lock del propio proyecto, no son una dependencia de registro sin pin. Un nombre sin prefijo (`app[dev]`) es para pip un requisito de registro y sigue siendo hallazgo, igual que una ruta local mezclada con un nombre sin pin o con `-r` sin `--require-hashes` (#980).
