# Alineamiento completo del gobierno CI

## Objetivo y alcance

El usuario solicita continuar paso a paso hasta alinear `DysTelefonica/access2web-blueprint` con el gobierno canónico y autoriza realizar los cambios necesarios. Prepare correcciones locales mínimas, verificadas y recuperables. No publique, mergee ni aplique etiquetas protegidas por esta autorización general. Los commits requieren una petición explícita conforme a la instrucción de seguridad vigente; no afirme cerradas las unidades de entrega sin ella.

El formato de las 36 issues abiertas ya está alineado y verificado en `odd/tasks/issue-format-alignment.md`; este documento sigue la configuración CI, la propagación y la evidencia de aceptación.

## Hallazgos y fuentes

- Consumer main sincronizado por fast-forward a `148c9165fdc57b9e5fb7cd1af97d16a3c03eb9c4`, sin sobrescribir `.pi/` ni `odd/` no versionados.
- Catálogo canónico auditado: `a40dcc3377327824a488ab4ed7f39ae174be50c6`, skill v1.14.10. El mirror global v0.4 está obsoleto y no sirve como autoridad de aceptación.
- CLI canónica `verify` devuelve exit 1: falta `manifest.parameters.P23_required_jobs` y falta `phases["6"].prs`.
- El registro de flota no declara `governance: ci-pattern`; el propagador selecciona la capa condicionalmente por ese campo, no por adoption.json.
- El CI del SHA actual tiene `required` en success, y las fases 0–5 pasan.
- Issue #782 abierta y con `status:approved`: aceptación de fase 6. No modifique su alcance ni afirme resultados ficticios.

## Tareas

- [x] G1 — Auditar CLI canónica, contratos, protección de main y propagación.
- [x] G2 — Añadir P23 al manifiesto del consumer, derivado de la política real. Implementación y verificación independiente aprobadas; commit y PR #828 publicados con autorización explícita, integración pendiente.
- [x] G3 — Preparar la declaración de gobierno del consumer en el registro de flota; comprobar el cambio y sus efectos sin propagación real. Commit y PR publicados; integración humana pendiente.
- [ ] G4 — Obtener evidencia verificable de un PR real para la fase 6; declarar solo resultados demostrados y comprobar el SHA en Git. BLOQUEADA: no se encontró un PR real integrado después de la fase 5.
- [ ] G5 — Verificar el conjunto final y distinguir cambios locales de cambios integrados en main. EN CURSO: CI remoto y entrega de P23 pendientes.
- [ ] G6 — Resolver la aprobación pendiente de #822. BLOQUEADA por ausencia de autoridad humana vinculada en el host; no condiciona la corrección técnica del arnés.

## Restricciones

- Un único escritor activo. Use worktrees dedicados y conserve todo trabajo ajeno.
- Primera unidad: únicamente `.governance-manifest.json` del consumer. Preserve `schema_version` y `files` existentes; no resele hashes sin cambios de contenido.
- G3 pertenece a `DysTelefonica/team-skills`; antes de editar allí lea las reglas actuales y el tour/propagador obligatorios en un checkout sincronizado separado. No cambie la rama de trabajo original del usuario.
- No implemente las issues de producto ni resuelva los datos pendientes de #771/#773/#774 o la decisión humana de #659.
- Los PRs #826/#827 de adopción no se aceptan automáticamente como ensayo de aceptación real. Los enlaces a checks de PRs distintos no prueban el éxito de un candidato.
- No instale dependencias ni cambie políticas, settings, runners, workflows o scripts para ocultar un rojo.
- No haga push, publicación de issues/PRs, merge, dispatch, aprobación ni propagación real sin autorización específica.

## Comprobaciones

G2: observe primero `parameters-absent` con la CLI canónica; después debe desaparecer, todos los jobs declarados deben estar cableados y los seis hashes existentes deben seguir verificados. El fallo esperado por fase 6 seguirá pendiente y se informará como tal, nunca como PASS global. Ejecute el validador propio del manifiesto del consumer si existe, con su runner disponible; no invente tests.

G3: compruebe JSON válido, un único cambio de registro, selección esperada del tier y dry-run sin escrituras externas. Declare runners o herramientas ausentes.

G4: confirme issue, rama, PR no artificial, checks del head exacto y merge correspondiente; haga `git cat-file -e <sha>^{commit}` y compruebe semántica además de estructura. Si no hay candidato suficiente, deje la fase 6 explícitamente pendiente.

G5: CLI canónica `verify` y `adoption check --all` con historial Git real; validación funcional/estructural propia del consumer y read-back de host al SHA pertinente. RDD solo según el interruptor del usuario y sobre una unidad Git concreta, no un TODO acumulado. No se afirma completitud antes de comprobar todas las salidas y la integración en main.

## Estado operativo

Worktree G2: `/home/ubuntu/repos/access2web-blueprint-worktrees/chore/782-governance-alignment`.
Rama G2: `chore/782-governance-alignment`, base `148c9165fdc57b9e5fb7cd1af97d16a3c03eb9c4`.
Documento y espejo Engram pertenecen al consumer. No versionados; preservados fuera del diff de producto del escritor.
Commits publicados: consumer `96d7c32fe26693850c53b071e5c22fe1813885df` en PR #828; catálogo `c865cffab4613f6de9ecb575db347809adf83244` en PR #397. No merge, no cierre de #782/#143 ni propagación real.

G2: escritor `mv1cp47q-8-x5j5` preparó 9 líneas aditivas en `.governance-manifest.json`, P23 igual a la política real. RED: `verify --json` exit 1 con `parameters-absent` y fase 6. GREEN parcial de unidad: exit 1 solo por fase 6, 5/5 jobs cableados y seis hashes conservados. `adoption check --phase 5` exit 0; `git diff --check` limpio. No existe suite local específica del manifiesto. El orquestador leyó el diff de una sola superficie y confirmó la adición. Assessment nativo: riesgo medium, RDD off, perfil fallback small; requiere verificador independiente, pendiente. No se marca G2 terminada hasta ese check.

G3 preparada para exploración: worktree sincronizado separado `/home/ubuntu/repos/team-skills-worktrees/access2web-governance`, rama `chore/143-access2web-governance`, base `a40dcc3377327824a488ab4ed7f39ae174be50c6`. Issue canónica existente #143 abierta y aprobada: graduar y distribuir ci-pattern a los consumers web. Ningún cambio en registro aún. La rama original `chore/preserve-cadete-349-taskfiles` sigue intacta.

Verificador `mv1dalt1-9-5a24`: G2 aprobado de forma independiente, P23 igual a la política, seis hashes recomputados y conservados, única superficie de 9 adiciones, fase 5 exit 0; el único fallo global restante en el consumer es fase 6. La lectura obligatoria de AGENTS/README/tour/propagador del checkout canónico sincronizado está completada. Propagación conserva P23: el resellado solo afecta las rutas que escribe el bot, no los parámetros ajenos. G3 requiere únicamente `fleet/registry.json`; issue #143 aprobada cubre esta distribución. No hay test determinista aplicable que produzca RED por esa declaración en este entorno; use checks estructurales y suites existentes sin inventar fixtures. `pwsh` ausente: propagador dry-run, doctor y pruebas PowerShell no ejecutables aquí. Es una comprobación pendiente, no un PASS.

G3: se añadió únicamente `governance: ci-pattern` a la entrada de `access2web-blueprint` en el worktree `chore/143-access2web-governance`. Assessment nativo: medium, RDD on, perfil runtime large; no exige verificador separado, pero las comprobaciones por comandos se delegan al verificador por la regla de ejecución. Falta observar la suite governance-partial y el validador frontmatter obligatorio, y el delta JSON exacto. Sin PowerShell no se simulará ni se afirmará el dry-run del propagador.

La activación YOLO del usuario concede permiso de commit/push no forzado/PR para cambios ordinarios ya acotados en el clon primario `team-skills`, sujeto a checks, RDD y política PR-only. No se transfiere automáticamente al clon independiente `access2web-blueprint` ni a los hijos; las restricciones de entrega de sus encargos siguen vigentes. No se hará merge automático.

G3 verificador `mv1dzops-a-q2ct`: delta JSON de un solo campo aprobado, governance-partial 10/10 exit 0 y frontmatter 63 archivos exit 0. Se mantiene pendiente el dry-run y doctor por ausencia de PowerShell. Review nativa `review-73c00a59ba4c289e`: una lente review-reliability, approved y acknowledgement completado con autoridad consumida para el candidato exacto. El consentimiento humano granted fue dirigido al worktree emisor; un primer intento de fachada desde la raíz principal fue rechazado antes de invocar native, sin linaje ni mutación, y se corrigió usando el workspaceRoot exacto.

Entrega G3 bajo permiso standing del clon primario: commit `c865cffab4613f6de9ecb575db347809adf83244`, `fix(fleet): declare ci governance for access2web`, push no forzado y PR `https://github.com/DysTelefonica/team-skills/pull/397`, base main, etiquetas type:chore y chain:partial, Refs #143 (no cierra la issue). Read-back confirmado: 1 adición, 0 eliminaciones, head exacto del commit. CI al consultar: smoke SUCCESS, unit suite IN_PROGRESS, nightly SKIPPED por diseño. No merge ni propagación real.

G4: consulta de PRs no detectó ninguno integrado después del merge de #827. #782 es una ISSUE de aceptación, no un PR; la conclusión del verificador que la llamaba ausente del repositorio se rechaza como confusión de tipo. No se registrará success ni evidencia de otro PR para pasar el gate.

Preferencia expresa del usuario: cambios mayores de 400 líneas (adiciones + eliminaciones) se entregan en PRs encadenados; intermedios chain:partial + Refs y cierre únicamente en la punta cuando complete el alcance. Las unidades actuales (P23 9 líneas, registro 1) no requieren división artificial.

Autorización explícita adicional del usuario: publicar commit, push no forzado y PR de P23 en `DysTelefonica/access2web-blueprint`, rama `chore/782-governance-alignment`, base main; sin merge, aprobación #822 o cierre de fase 6. Se ejecutó commit `96d7c32fe26693850c53b071e5c22fe1813885df`, `fix(governance): declare required jobs in manifest`, y publicación `https://github.com/DysTelefonica/access2web-blueprint/pull/828`, etiquetas type:chore/chain:partial, Refs #782. Read-back confirma 9 adiciones/0 eliminaciones y head exacto. CI al consultar: review-budget, gitleaks, pip-audit y codeql en curso; quality y trivy-config en cola; mutation skipped. Ningún PASS remoto global afirmado. El push informó 12 alertas Dependabot preexistentes en main (4 high, 6 moderate, 2 low); no se han auditado ni corregido en esta unidad.

## Siguiente paso

CI remoto de PRs #397 y #828 pendiente; deje que lo resuelva el mecanismo del host, sin bucles de polling y sin merge automático. El mantenedor integrará solo tras CI verde y revisión. Fase 6 necesita todavía un PR real de trabajo posterior a la adopción, no estas correcciones del contrato; no declare success ficticio. Dry-run/doctor con PowerShell y aprobación #822 siguen pendientes explícitos.
