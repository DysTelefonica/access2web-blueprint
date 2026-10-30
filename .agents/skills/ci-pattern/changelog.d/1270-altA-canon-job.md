type: minor

#1270 (alternativa A): el secret del catálogo sale de los jobs de PR. Los gates pasan a un workflow en `pull_request_target`, que ejecuta el workflow de la base: ahí el job `canon` hace checkout de la BASE (de donde salen la acción y el sello), usa el token del catálogo en un solo paso y publica el patrón como artefacto tarball; los jobs del gate lo consumen sin ningún secreto y leen la cabeza del PR como datos.

El sello se lee de la BASE, así que un PR no puede apuntar a un commit más antiguo del catálogo: el residuo del "sello antiguo" desaparece. El patrón se mide contra el commit que la base declara.

El veredicto se publica como estado de commit sobre la cabeza del PR: un run de `pull_request_target` se atribuye a la base y su check run no bloquea el merge.

El canon añade la regla que lo hace cumplir: un job que puede correr en `pull_request` no puede referenciar `secrets.TEAM_SKILLS_READ_TOKEN`. La regla sigue a los workflows reutilizables hasta sus llamadores, porque heredan los eventos de quien los llama.
