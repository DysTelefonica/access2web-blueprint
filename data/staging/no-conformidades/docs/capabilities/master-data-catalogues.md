# Capacidad: maestros y catÃ¡logos

## Â§0 Identidad
- **ID de capacidad**: `CAP-MASTER-CATALOGUES`
- **Tier**: standard
- **Estado**: active / inventario documental inicial
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmaciÃ³n â€” Calidad / administraciÃ³n de datos maestros
- **Ãšltima verificaciÃ³n**: 2026-06-15 mediante inspecciÃ³n estÃ¡tica; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixta â€” algunos catÃ¡logos tienen pruebas registradas; otros solo nombres/cÃ³digo

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Mantener vocabularios controlados para tipologÃ­as, motivos de control de eficacia, estados, tÃ©cnicos, responsables, jurÃ­dicas y otros datos maestros.
- **Usuarios / perfiles**: Administradores/calidad, responsables de dominio, usuarios que filtran/listan NC.
- **Problema que resuelve**: Evita valores libres incompatibles y filtros rotos en formularios, indicadores e informes.
- **Valor de negocio / por quÃ© existe**: Los catÃ¡logos son reglas de negocio: determinan clasificaciÃ³n, visibilidad, estados y controles.
- **No-objetivos**: No documenta el ciclo de vida completo de cada NC.
- **Origen de la intenciÃ³n**: Formularios/clases `Tipologia*`, `FormMotivosNoRequiereControlEficacia`, bootstrap de estado y tests issue-3/47.
- **Referencia de tracker de origen**: Issues #3, #24, #47, #67.

## Â§2 Contrato de comportamiento

### Escenarios (Dado / Cuando / Entonces)
- **DADO** que un usuario registra motivo de â€œno requiere control de eficaciaâ€ **CUANDO** confirma el formulario **ENTONCES** se emite `MotivoRegistrado` con el texto actual.
- **DADO** que se inicializa el catÃ¡logo de estados **CUANDO** faltan cÃ³digos esperados **ENTONCES** la carga debe fallar de forma explÃ­cita.
- **DADO** que se gestiona tipologÃ­a de NC Proyecto **CUANDO** se abre desde menÃº de configuraciÃ³n **ENTONCES** solo usuarios autorizados acceden.
- **DADO** que se normalizan responsables o tipologÃ­as **CUANDO** se ejecutan operaciones de instalador/migraciÃ³n **ENTONCES** los valores resultantes deben ser idempotentes y verificables.

### Reglas de negocio
| ID regla | Enunciado (pretendido) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-CAT-1 | El motivo de no requerir control de eficacia tiene campos de dominio y persiste en NC Proyecto y NC AuditorÃ­a. | Tests issue-3 | SÃ­ segÃºn manifest principal | `Test_MotivoNoRequiereControlEficacia_DomainFields_Atomic`, `Test_E2E_MotivoPersistencia_*`; no reejecutado | Verified-static |
| BR-CAT-2 | El formulario de motivos emite evento `MotivoRegistrado` con el valor actual. | CÃ³digo exportado | SÃ­ â€” `Form_FormMotivosNoRequiereControlEficacia` | FALTA â†’ author via access-vba-tdd de contrato de formulario/costura | Verified-static |
| BR-CAT-3 | El catÃ¡logo de estados se crea/idempotente y conserva cÃ³digos esperados. | Tests issue-47 | SÃ­ segÃºn manifests | `Test_EstadoCatalogo_*`; no reejecutado | Verified-static |
| BR-CAT-4 | La recarga de diccionario de estados falla si falta un cÃ³digo esperado. | Tests issue-47 | SÃ­ segÃºn manifest principal | `Test_EstadoCatalogo_DictionaryReload_FailsOnMissingCode_Atomic`; no reejecutado | Verified-static |
| BR-CAT-5 | La gestiÃ³n de tipologÃ­a NC Proyecto estÃ¡ restringida a no tÃ©cnicos/autorizados. | CÃ³digo exportado | Parcial â€” menÃº bloquea tÃ©cnicos para configuraciÃ³n | FALTA â†’ author via access-vba-tdd | Verified-static |
| BR-CAT-6 `#79 (BR-CAT-6/7)` | CatÃ¡logos de tÃ©cnicos, responsables, jurÃ­dicas, proveedores y tipologÃ­as tienen contrato de alta/baja/ediciÃ³n y uso en filtros. | Producto pendiente | Desconocido/parcial por nombres | FALTA â†’ author via access-vba-tdd tras inventario de esquema. Cross-link: misma matriz de cobertura referenciada por `users-permissions-navigation` BR-UPN-7 y `cross-cutting-support` BR-XCUT-6 | Intended |
| BR-CAT-7 `#79 (BR-CAT-6/7)` | Normalizaciones de instalador para responsables/tipologÃ­as son idempotentes y no destruyen valores vÃ¡lidos. | CÃ³digo exportado | Probable â€” `Instalador` tiene rutinas | FALTA â†’ author via access-vba-tdd | Intended |

### Validaciones
- No aÃ±adir motivos vacÃ­os si producto lo prohÃ­be â€” pendiente de confirmaciÃ³n.
- Estados esperados deben existir antes de confiar en cachÃ©s/listados.
- La ediciÃ³n de catÃ¡logos sensibles requiere permisos explÃ­citos.

### Transiciones de estado
- `CatÃ¡logo ausente` --(`BootstrapEstadoCatalogo`)--> `CatÃ¡logo creado`.
- `CatÃ¡logo existente` --(`BootstrapEstadoCatalogo`)--> `Sin duplicados / idempotente`.
- `Motivo escrito` --(`ComandoAceptar`)--> `MotivoRegistrado`.

### Casos lÃ­mite y de error
- CatÃ¡logo de estados incompleto debe fallar rÃ¡pido, no degradar silenciosamente.
- Los catÃ¡logos con nombres parecidos (`Responsable`, `ResponsableCalidad`, `TÃ©cnico`) requieren esquema primero para evitar falsos positivos.

### SeÃ±ales de aceptaciÃ³n / presencia
- Tests de motivo y estado pasan en staging actual.
- Cada catÃ¡logo usado por filtros/listas tiene pÃ¡gina o secciÃ³n con reglas, permisos y pruebas.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada de UI**: `Form_FormNCProyectoTipologiaGestion`, `Form_FormTipologiaNCProyecto`, `Form_FormMotivosNoRequiereControlEficacia`, `Form_Form0BDTecnicos`, menÃºs de configuraciÃ³n.
- **Puntos de entrada de cÃ³digo**: `TipologiaNCProyectos`, `EstadoCatalogoBootstrap`, `Instalador`, `Entorno` colecciones `ColTipos`, `ColEstadosNC`, `ColUsuariosCalidad`, `ColJefesProyecto`, `Juridica`.
- **Datos afectados**: `TbTipologia`, tablas de estados, motivos no requiere CE, usuarios/tÃ©cnicos/responsables/jurÃ­dicas/proveedores exactos pendientes de esquema.
- **Salidas**: combos, filtros, captions, validaciones de dominio.
- **Dependencias e integraciones**: control eficacia, listados, indicadores, expediente/responsables, usuarios/permisos.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada; tarea solo documental.
- **ValoraciÃ³n de diseÃ±o**: los catÃ¡logos estÃ¡n dispersos entre formulario, entorno e instalador. Para migraciÃ³n conviene centralizarlos como APIs de catÃ¡logo con permisos y versionado.

## Â§4 Receta de reconstrucciÃ³n
1. Inventariar tablas reales de catÃ¡logo y sus FK antes de escribir fixtures.
2. Separar catÃ¡logos crÃ­ticos: estados, tipologÃ­as, motivos CE, usuarios/tÃ©cnicos/responsables, jurÃ­dicas/proveedores.
3. Crear pruebas de bootstrap/idempotencia/validaciÃ³n por catÃ¡logo.
4. Crear pruebas de UI/costura para formularios de gestiÃ³n de catÃ¡logo solo donde aporten contrato de negocio.

## Â§5 Evidencia y trazabilidad
- **Tests**: `tests/tests.vba.json` contiene pruebas de motivo no requiere CE y estado catÃ¡logo; `tests/tests.vba.cache-readiness.json` contiene warm-up de catÃ¡logo de estados. No se reejecutaron.

| Elemento | Ref. tracker | VersiÃ³n de staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en producciÃ³n | Nota |
|---|---|---|---|---|---|---|
| Motivo no requiere CE | Issue #3 | Pendiente | pending | Pendiente | Pendiente | Tests registrados. |
| Icono/botÃ³n motivos NR | Issue #24 | Pendiente | pending | Pendiente | Pendiente | Evidencia de issue closeout; capacidad funcional necesita prueba actual. |
| CatÃ¡logo de estados | Issue #47 | Pendiente | pending | Pendiente | Pendiente | Tests registrados. |
| TipologÃ­as/tÃ©cnicos/proveedores/responsables | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta inventario/pruebas. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla del documento |
|---|---|---|---|
| Estado desaparece de filtros/listados | CatÃ¡logo de estados incompleto | Reejecutar issue-47 | BR-CAT-3..4 |
| Motivo CE no persiste | RegresiÃ³n de motivo no requiere CE | Reejecutar issue-3 | BR-CAT-1..2 |
| TipologÃ­a no editable o filtro roto | CatÃ¡logo sin contrato | Crear prueba de tipologÃ­a | BR-CAT-5..6 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- El motivo de "no requiere control de eficacia" persiste en NC Proyecto y NC AuditorÃ­a con sus campos de dominio (BR-CAT-1): la web debe seguir garantizando que un motivo registrado desde `Form_FormMotivosNoRequiereControlEficacia` se persista y se pueda consultar desde ambas trayectorias de NC.
- El evento `MotivoRegistrado` se emite con el valor actual del motivo al confirmar (BR-CAT-2): la web debe exponer un endpoint que registre el motivo y devuelva el evento/movimiento al consumidor, replicando el `ComandoAceptar_Click` del formulario.
- El catÃ¡logo de estados se crea/idempotente y conserva los cÃ³digos esperados (BR-CAT-3): `BootstrapEstadoCatalogo` debe seguir presente en el servicio de bootstrap; la web debe poder llamarlo y verificar que los cÃ³digos esperados estÃ¡n, sin duplicar.
- La recarga del diccionario de estados falla si falta un cÃ³digo esperado (BR-CAT-4): la web debe propagar el error con el cÃ³digo ausente, no degradar silenciosamente ni continuar con un catÃ¡logo incompleto.
- La gestiÃ³n de tipologÃ­a de NC Proyecto estÃ¡ restringida a no tÃ©cnicos/autorizados (BR-CAT-5): la API web debe chequear el rol y devolver `403` para tÃ©cnicos, replicando el bloqueo de menÃº de `Form_Form0BDOpcionesParteProyectos.cls`.
- Los catÃ¡logos de tÃ©cnicos, responsables, jurÃ­dicas, proveedores y tipologÃ­as exponen contrato de alta/baja/ediciÃ³n y se usan en filtros (BR-CAT-6): la web debe permitir CRUD en cada catÃ¡logo, con permisos diferenciados, y mantener la integridad referencial con NC.
- Las normalizaciones de instalador para responsables/tipologÃ­as son idempotentes (BR-CAT-7): una migraciÃ³n que se ejecute dos veces sobre el mismo dataset no debe duplicar filas ni destruir valores vÃ¡lidos.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir los formularios `Form_FormNCProyectoTipologiaGestion`, `Form_FormTipologiaNCProyecto`, `Form_FormMotivosNoRequiereControlEficacia`, `Form_Form0BDTecnicos` por endpoints REST `GET/POST/PUT/DELETE` con autenticaciÃ³n, autorizaciÃ³n y versionado, no por formularios Access.
- Convertir `EstadoCatalogoBootstrap` en un job de bootstrap del backend con un endpoint `POST /catalogos/estados/bootstrap` que valide los cÃ³digos esperados, no por un mÃ³dulo VBA ejecutado al abrir la app.
- Reemplazar el patrÃ³n "combos poblados desde `Entorno.Col*`" por catÃ¡logos servidos desde la API, cacheables en cliente con TTL, en lugar de colecciones globales en memoria.
- Mover la instalaciÃ³n/normalizaciÃ³n de catÃ¡logos (`Instalador`) a migraciones versionadas de base de datos con `up`/`down` explÃ­citos, no por rutinas in-process disparadas por eventos.
- Sustituir el conjunto disperso de catÃ¡logos (tipologÃ­as, motivos CE, tÃ©cnicos, responsables, jurÃ­dicas, proveedores) por un Ãºnico servicio de catÃ¡logos con discriminador `tipoCatalogo`, no por cinco clases paralelas.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar la normalizaciÃ³n puntual del instalador como regla permanente: la web debe distinguir entre "migraciÃ³n inicial" (idempotente) y "regla de negocio" (continua), y no mezclarlas.
- No usar la visibilidad de un menÃº como control de seguridad real para catÃ¡logos sensibles: la web debe aplicar permisos en el servidor, no en la UI.
- No exponer catÃ¡logos como colecciones en memoria (`ColTipos`, `ColEstadosNC`, `ColUsuariosCalidad`): la web debe servirlos desde base de datos o cachÃ© con TTL explÃ­cito, no mantenerlos como globales de proceso.
- No migrar la combinaciÃ³n `Entorno` + `Instalador` como Ãºnica puerta de entrada: la web debe tener un servicio de catÃ¡logos con endpoints CRUD para cada tipo.
- No portar la duplicaciÃ³n entre "motivoCE" en `TbMotivosNoRequiereControlEficacia` y "tipologÃ­a" en `TbTipologia` con la misma UI: la web debe tratarlos como catÃ¡logos independientes con UIs dedicadas, no como un Ãºnico formulario genÃ©rico.

### Â§6.4 Preguntas abiertas al product owner
- Â¿El catÃ¡logo de estados es el mismo para Proyecto y AuditorÃ­a o se diferencia? (BR-CAT-3) Confirmar lista canÃ³nica de estados por dominio.
- Â¿Los motivos de "no requiere control de eficacia" son los mismos para NC Proyecto y NC AuditorÃ­a? (BR-CAT-1) Â¿O cada dominio tiene su propio subconjunto?
- Â¿La gestiÃ³n de tipologÃ­a NC Proyecto (BR-CAT-5) es por Calidad, por un rol especÃ­fico, o por un permiso por proyecto? Confirmar la regla de rol.
- Â¿Los catÃ¡logos de tÃ©cnicos, responsables, jurÃ­dicas, proveedores y tipologÃ­as (BR-CAT-6) requieren workflow de aprobaciÃ³n o el alta/baja es directa?
- Â¿La normalizaciÃ³n del instalador (BR-CAT-7) tiene una fecha de corte o se mantiene corriendo en cada release? Hoy se ejecuta como parte de `Instalador`; Â¿se mantiene ese patrÃ³n o se eliminarÃ¡ tras la primera migraciÃ³n?
- Â¿QuÃ© versiÃ³n de cada catÃ¡logo debe quedar congelada cuando una NC se cierra? Confirmar si las NC cerradas deben "ver" el catÃ¡logo vigente al cierre o el actual.

## Â§7 Registro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-CAT-1 â€” El motivo de no requerir control de eficacia tiene campos de dominio y persiste en NC Proyecto y NC AuditorÃ­a. | Verified-static | `Test_MotivoNoRequiereControlEficacia_DomainFields_Atomic`, `Test_E2E_MotivoPersistencia_*`; FALTA â†’ reejecutar | 2026-06-15 |
| BR-CAT-2 â€” El formulario de motivos emite evento `MotivoRegistrado` con el valor actual. | Verified-static | `Form_FormMotivosNoRequiereControlEficacia`; FALTA â†’ author via access-vba-tdd de contrato de formulario/costura | 2026-06-15 |
| BR-CAT-3 â€” El catÃ¡logo de estados se crea/idempotente y conserva cÃ³digos esperados. | Verified-static | Familia `Test_EstadoCatalogo_*` registrada en `tests/tests.vba.json`; FALTA â†’ reejecutar | 2026-06-15 |
| BR-CAT-4 â€” La recarga de diccionario de estados falla si falta un cÃ³digo esperado. | Verified-static | `Test_EstadoCatalogo_DictionaryReload_FailsOnMissingCode_Atomic`; FALTA â†’ reejecutar | 2026-06-15 |
| BR-CAT-5 â€” La gestiÃ³n de tipologÃ­a NC Proyecto estÃ¡ restringida a no tÃ©cnicos/autorizados. | Verified-static | MenÃº bloquea tÃ©cnicos para configuraciÃ³n; FALTA â†’ author via access-vba-tdd | 2026-06-15 |
| BR-CAT-6 â€” CatÃ¡logos de tÃ©cnicos, responsables, jurÃ­dicas, proveedores y tipologÃ­as tienen contrato de alta/baja/ediciÃ³n y uso en filtros. | Intended | FALTA â†’ author via access-vba-tdd tras inventario de esquema; cross-link `users-permissions-navigation` BR-UPN-7 y `cross-cutting-support` BR-XCUT-6 | 2026-06-15 |
| BR-CAT-7 â€” Normalizaciones de instalador para responsables/tipologÃ­as son idempotentes y no destruyen valores vÃ¡lidos. | Intended | FALTA â†’ author via access-vba-tdd | 2026-06-15 |
| Hay tests registrados para motivo no requiere CE. | Verified-static | `tests/tests.vba.json` | 2026-06-15 |
| Hay tests registrados para catÃ¡logo de estados. | Verified-static | `tests/tests.vba.json`, `tests/tests.vba.cache-readiness.json` | 2026-06-15 |
| Existe formulario de motivos que emite evento con el motivo. | Verified-static | `src/forms/Form_FormMotivosNoRequiereControlEficacia.cls` | 2026-06-15 |
| CatÃ¡logos de proveedores/tÃ©cnicos/responsables estÃ¡n completamente especificados. | Intended | Falta inventario de esquema y pruebas | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia confirmada. Hueco: hay catÃ¡logos visibles por nombres y tests parciales, pero falta contrato completo por catÃ¡logo.
