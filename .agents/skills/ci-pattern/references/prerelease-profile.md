# Perfil pre-release de ci-pattern (#1627)

Referencia completa del perfil. La versión corta que leen los agentes está en
`slices/partials/web.md`, sección «Perfil pre-release». Esta página es la fuente
de detalle; no la copie a otros sitios.

## Ámbito y fin

- Rige todo proyecto con `primary_type: web`.
- Se activa con un fichero versionado en el repo del proyecto, `.github/prerelease.json`,
  con dos campos: `enabled` (booleano, obligatorio) y `schema_version` (opcional, entero
  >= 1; si falta vale `1`). El validador del catálogo acepta el fichero actual de APAP_WEB
  `.github/prerelease.json`, `{"enabled": true}`. APAP_WEB no debe añadir `schema_version` ni BOM
  mientras su propio `scripts/check_prerelease.py` solo acepte `enabled`.
- Validador: `assets/prerelease/check_prerelease_profile.py <ruta>`. Sale con `0` si la
  declaración es válida y con `2` nombrando cada campo inválido. Un fichero ausente,
  un JSON no válido o un campo desconocido también salen con `2`. Las claves duplicadas,
  un fichero que no es UTF-8 y un fichero UTF-16 salen con `2` sin traceback; un BOM
  UTF-8 se acepta.
- Termina en la primera release del proyecto. Después aplica el perfil normal
  (decisión del dueño).

## Reglas

1. **Día, cada PR ordinario.** Solo gates baratos: lint, typecheck, tests unitarios,
   pr-gate, los ficheros e2e propios del PR y la seguridad de dependencias si cambian
   las dependencias. Integración y ruta de actualización solo si el PR toca SQL o
   migraciones. Se fusiona en cuanto está en verde: sin deploy, sin job caro y sin
   re-ejecución tras la fusión.
2. **Noche, una canalización sobre la cabeza de main.** Construye la imagen y ejecuta
   todos los checks caros: e2e completo, integración, ruta de actualización, Trivy y
   escaneo de seguridad profundo. Despliega a producción solo si todos están en verde
   sobre el mismo SHA.
3. **Noche en rojo.** Sin deploy. Una issue de máxima prioridad por cada job en rojo,
   por delante de cualquier trabajo nuevo. La canalización nocturna se vuelve a lanzar
   a mano el mismo día, en cuanto esté en verde. Objetivo: al menos un deploy exitoso
   al día.
4. **Migraciones y DDL.** El deploy espera siempre a una copia de seguridad verificada.
   La canalización nocturna nunca aplica sola una migración pendiente.
5. **Cadenas.** Como máximo tres eslabones abiertos; cada uno se fusiona en cuanto está
   en verde. Un script abre cada eslabón con la etiqueta y `Refs #N`; solo la punta lleva
   `Closes #N`.
6. **Nada redundante.** Ningún check corre dos veces sobre el mismo código el mismo día
   (PR, luego push a main, luego cron). Si es caro, corre de noche.
7. **Fin del perfil.** Ver «Ámbito y fin».

Reglas adicionales del dueño:

- Todo paso con una regla fija es un script versionado y con tests, nunca un turno del
  modelo.
- Toda automatización sale con sus tests escritos antes que el código, y el `.sh` del
  runner tiene modo de ensayo.
- Una noche en rojo se vuelve a lanzar el mismo día hasta que el deploy a producción
  salga bien.
- El runner nocturno se aprovisiona con un `.sh` idempotente que vive en team-skills o
  en infra, nunca en APAP_WEB.
- Todo PR que toque la canalización nocturna hace una ejecución nocturna de ensayo antes
  de fusionarse.

## Regla dura del dueño

Mientras la canalización nocturna esté en rojo, no se trabaja ni se fusiona ninguna
issue salvo las que arreglan ese rojo. La línea se para hasta que la ejecución nocturna
salga en verde y se despliegue. Rige ya desde ahora; además, el merge gobernado la hará cumplir de forma automática.

## Relajaciones de cadena y re-ejecución

Las reglas de contexto de cadena (tabla de contexto obligatoria) y de re-ejecución de CI
tras un retarget se relajan **solo** bajo el perfil activo: no se exige la tabla, y no
se vuelve a ejecutar CI tras un retarget cuando la cabeza no ha cambiado. Fuera del
perfil rigen sin cambios. Decisión registrada en la issue #1627.

## Estado de la implementación

Este PR cubre: el esquema versionado, el validador con sus tests, la sección del partial
web y esta referencia. No cubre todavía los criterios de ejecución (ningún check caro en
un PR ordinario, deploy nocturno solo con todo en verde). Esos llegan con los slices
siguientes.

## Slices posteriores (propuesta de corte)

Cada automatización sale con sus tests escritos antes que el código, y con modo de
ensayo cuando toca la canalización.

| Slice | Contenido | Criterios de #1627 que cierra |
|---|---|---|
| 2 | Gate de parada de línea en el merge gobernado: lee solo la etiqueta fija y la conclusión de la última ejecución nocturna; falla cerrado si no puede leerla. | Regla dura; merge rechaza PRs sin issue de arreglo |
| 3 | Canalización nocturna sobre la cabeza de main, con modo de ensayo, más el script de re-lanzamiento. | Deploy nocturno solo con todo en verde sobre el mismo SHA |
| 4 | Informe nocturno en HTML: PRs desde la última noche buena, resultado de cada job caro, si hubo deploy. | Script de informe con tests |
| 5 | Aparcar y desaparcar las issues normales durante una noche en rojo. | Script de aparcar/desaparcar con tests |
| 6 | `.sh` idempotente que aprovisiona el runner nocturno, con modo de ensayo; sale con `2` si algo no cuadra. Vive en team-skills o infra, nunca en APAP_WEB. | `.sh` idempotente con ensayo y test |
| 7 | Script que abre cada eslabón de cadena; comprobación diaria de deploy. | Cadenas automatizadas; un deploy al día |
| 8 | Ensayo nocturno obligatorio antes del merge de todo PR que toque la canalización nocturna, y tests de no-coste en PR ordinario. | Ningún check caro en PR ordinario |

`SKILL.md` de ci-pattern no recibe líneas en este slice: está en su tope de 1000.

## Instantánea de la noche (slice 4b)

`build_night_snapshot.py --repo OWNER/NAME --out snap.json` (o `--dry-run`) construye la
instantánea de `night_report.py` con `gh api` en solo lectura: PRs desde la última noche
buena con sus issues, jobs caros de la noche y deploy a `production` del día UTC.
