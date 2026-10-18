type: minor

El gate `pr-size` admite `--event <pull_request event json>`: cuenta las líneas con la API de GitHub (`GITHUB_TOKEN`; `GITHUB_API_URL` solo https, o http en loopback) sin necesitar historial de git, lee las etiquetas del propio evento y falla cerrado ante cualquier error de la API o si el desglose de ficheros no coincide con `changed_files` del PR. Aplica las mismas reglas de raíces vendorizadas y de renombrados que el modo git (#806).
