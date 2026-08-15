# Expedientes — catálogo canónico para la migración web

## Decisión que organiza el catálogo

La migración conserva **capacidades y resultados de negocio**, no formularios ni mecanismos Access. Este catálogo sustituye el índice legacy como contrato de alcance para las specs de `expedientes-web-migration`: ninguna entrada puede desaparecer sin una disposición nueva aprobada.

La épica histórica [`epic.md`](epic.md) queda superseded como plan de producto porque prescribe React y paridad de 46 pantallas. Sigue siendo evidencia de inventario, no arquitectura objetivo.

## Cómo leer la evidencia y la disposición

| Marca | Significado |
|---|---|
| `VR-hist` | Existía evidencia runtime en el catálogo legacy; no se ha repetido en esta fase. |
| `VS` | Evidencia estática en source/documentación auditados. |
| `INT` | Contrato pretendido o integración pendiente de confirmar. |
| `PRESERVAR` | Mantener el resultado y sus reglas observadas. |
| `MODERNIZAR` | Mantener el resultado sustituyendo el mecanismo legacy. |
| `RETIRAR` | Eliminar el mecanismo, sin retirar silenciosamente un resultado de negocio. |

**Límite de evidencia:** el inventario actual confirma 46 formularios, 143 módulos BAS (70 de producción y 73 de pruebas/harness), 66 clases CLS (65 de producción y un DTO de test) y 260/260 átomos válidos en el manifest. Las 49 tablas fueron observadas en runtime el 2026-08-05; esta fase no repitió esa validación, por lo que su schema, uso y volumen deben refrescarse antes del DDL final. Dysflow queda como evidencia de solo lectura sobre `expedientes-e2e-integration-rest`; el desajuste de configuración impide usarlo como prueba runtime actual de esta fase.

## Corrección del linaje legacy

- La antigua `CAP-045` nombraba dos capacidades: **basados/lotes** pasa a `EXP-CAP-006` y **jurídicas** queda incluida en `EXP-CAP-013`.
- Las antiguas `CAP-046` y `CAP-047` describían el mismo ordinal E2E desde dos capas; se consolidan en `EXP-CAP-042` sin perder reglas ni trazabilidad.
- Las 56 especificaciones legacy quedan cubiertas por referencias explícitas en la columna «Linaje»; el nuevo ID es el único identificador normativo.

## Ledger normativo

### A. Ciclo de vida

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-001 | Alta transaccional del expediente, incluida el alta iniciada por sistemas autorizados. | CAP-001 | VS | PRESERVAR |
| EXP-CAP-002 | Edición del agregado con invariantes y concurrencia explícitas. | CAP-002 | VS | PRESERVAR |
| EXP-CAP-003 | Baja condicionada, auditable y sin pérdida accidental de relacionados. | CAP-003 | VS | PRESERVAR |
| EXP-CAP-004 | Cambio de tipo con validación de elegibilidad y efectos derivados. | CAP-004 | VS | PRESERVAR |
| EXP-CAP-005 | Estado calculado y transiciones con fechas, garantía y motivos. | CAP-005 | VS | PRESERVAR |
| EXP-CAP-006 | Jerarquía de acuerdos marco, lotes y basados, incluida integridad padre-hijo. | CAP-045 (basados/lotes) | VS | PRESERVAR |
| EXP-CAP-007 | Derivación ordinal funcional dentro de la jerarquía. | CAP-046 (parte no E2E) | VR-hist | PRESERVAR |

### B. Datos relacionados

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-008 | Hitos y sus reglas temporales. | CAP-016 | VS | PRESERVAR |
| EXP-CAP-009 | Modificados, cambios de campo e historial reciente. | CAP-017, CAP-039 | VS | PRESERVAR |
| EXP-CAP-010 | Anexos y referencias documentales con eliminación y retención controladas. | CAP-018 | VS | MODERNIZAR |
| EXP-CAP-011 | Anualidades vinculadas al expediente. | CAP-043 | VS | PRESERVAR |
| EXP-CAP-012 | Responsables por rol y jefaturas. | CAP-044 + inventario | VS | PRESERVAR |
| EXP-CAP-013 | Relaciones con entidades, jurídicas y cadena de entidades. | CAP-045 (jurídicas), CAP-052 | VS | PRESERVAR |
| EXP-CAP-014 | Suministradores, contratistas, subcontratistas, UTE y jerarquía «el árbol manda». | CAP-019 | VS | MODERNIZAR |

### C. Diez catálogos

| ID | Catálogo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-015 | Comerciales. | CAP-006 | VS | PRESERVAR |
| EXP-CAP-016 | CPV. | CAP-007 | VS | PRESERVAR |
| EXP-CAP-017 | Ejércitos. | CAP-008 | VS | PRESERVAR |
| EXP-CAP-018 | Suministradores. | CAP-009 | VS | PRESERVAR |
| EXP-CAP-019 | Lugares de ejecución. | CAP-010 | VR-hist/VS | PRESERVAR |
| EXP-CAP-020 | PECAL. | CAP-011 | VS | PRESERVAR |
| EXP-CAP-021 | RAC. | CAP-012 | VR-hist/VS | PRESERVAR |
| EXP-CAP-022 | Grados de clasificación. | CAP-013 | VS | PRESERVAR |
| EXP-CAP-023 | Órganos de contratación. | CAP-014 | VS | PRESERVAR |
| EXP-CAP-024 | Oficinas de programa. | CAP-015 | VS | PRESERVAR |

### D. Consulta, bandejas y tareas

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-025 | Bandeja paginada con filtros y detalle de solo lectura. | CAP-022 | VS | MODERNIZAR |
| EXP-CAP-026 | Búsqueda avanzada con semántica de filtros preservada. | CAP-023 | VR-hist/VS | PRESERVAR |
| EXP-CAP-027 | Búsqueda técnica sujeta a permiso efectivo. | CAP-024 | VS | PRESERVAR |
| EXP-CAP-028 | Exportación Excel reproducible y autorizada. | CAP-025 | VS | PRESERVAR |
| EXP-CAP-029 | Tareas calculadas por reglas del expediente, sin inventar una tabla de tareas. | CAP-040 | INT/VS | MODERNIZAR |

### E. Escritura resiliente y feedback

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-030 | Autosave de datos generales y fechas con transacción explícita. | CAP-020 | VR-hist | MODERNIZAR |
| EXP-CAP-031 | Autosave de hitos, modificados y restantes secciones editables. | CAP-021 | VS | MODERNIZAR |
| EXP-CAP-032 | Idempotencia, prevención de doble envío y feedback accesible de progreso/error. | CAP-035, CAP-036 | VS | MODERNIZAR |

### F. E2E, JSON y trazabilidad

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-033 | DTO de dominio estable para intercambio. | CAP-038 | VS | PRESERVAR |
| EXP-CAP-034 | Exportación JSON completa y determinista. | CAP-029 | VR-hist | PRESERVAR |
| EXP-CAP-035 | Gestión y ejecución batch E2E. | CAP-026 | VR-hist | MODERNIZAR |
| EXP-CAP-036 | Hash de contenido con algoritmo y canonicalización versionados. | CAP-027 | VR-hist | PRESERVAR |
| EXP-CAP-037 | Generación del paquete de exportación E2E. | CAP-028 | VR-hist | PRESERVAR |
| EXP-CAP-038 | Trazabilidad entre expediente, batch, artefacto y resultado. | CAP-030 | VR-hist | PRESERVAR |
| EXP-CAP-039 | Selección manual de expedientes para batch. | CAP-048 | VR-hist | PRESERVAR |
| EXP-CAP-040 | Destino de exportación por usuario, sin rutas locales como contrato. | CAP-049 | VR-hist | MODERNIZAR |
| EXP-CAP-041 | Sesión de exportación, reanudación y cierre. | CAP-050 | VR-hist | PRESERVAR |
| EXP-CAP-042 | Ordinal E2E único y consistente en dominio y trazabilidad. | CAP-046 + CAP-047 | VR-hist | PRESERVAR |

### G. Identidad, autorización y auditoría

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-043 | Autorización deny-by-default por capacidad. | CAP-041 | VS | MODERNIZAR |
| EXP-CAP-044 | Login y principal actual consumidos desde Lanzadera. | CAP-042 | VS | MODERNIZAR |
| EXP-CAP-045 | Estado de usuario y sesión sin globals mutables. | CAP-053 | VS | MODERNIZAR |
| EXP-CAP-046 | Auditoría de accesos, cambios, exportaciones e integraciones. | CAP-039 + inventario | VS | MODERNIZAR |

El único contrato de identidad objetivo es `CurrentPrincipal/effective_permissions`. Cualquier permiso ausente, desconocido o no resoluble deniega la operación.

### H. Runtime, configuración y readiness

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-047 | Arranque y readiness con fallos diagnosticables. | CAP-032 | VR-hist | MODERNIZAR |
| EXP-CAP-048 | Configuración tipada por entorno e inyección de dependencias. | CAP-037 | VS | MODERNIZAR |
| EXP-CAP-049 | Caché invalidable y read-models reconstruibles. | CAP-033 | VR-hist | MODERNIZAR |
| EXP-CAP-050 | Binding de backend por despliegue, sin selector interactivo ni rutas embebidas. | CAP-031 | VR-hist | MODERNIZAR |

### I. Integraciones

| ID | Capacidad y obligación objetivo | Linaje | Evidencia | Disposición |
|---|---|---|---|---|
| EXP-CAP-051 | Alta/consulta desde HPS mediante contrato autenticado e idempotente. | inventario | VS/INT | MODERNIZAR |
| EXP-CAP-052 | Intercambio con AGEDYS con ownership y errores definidos. | CAP-051 | INT | PRESERVAR |
| EXP-CAP-053 | Enlace con Gestión de Riesgos por identificador estable. | inventario | VS/INT | PRESERVAR |
| EXP-CAP-054 | Enlace con No Conformidades y su estado asociado. | CAP-034 | VR-hist/INT | PRESERVAR |
| EXP-CAP-055 | Notificaciones por correo mediante puerto y trazabilidad. | inventario | VS/INT | MODERNIZAR |
| EXP-CAP-056 | Documentos en SharePoint u object storage mediante referencia opaca, retención y autorización. | CAP-018 + inventario | VS/INT | MODERNIZAR |

### J. Mecanismos legacy retirados o reemplazados

| ID | Mecanismo | Linaje | Sustitución verificable | Disposición |
|---|---|---|---|---|
| EXP-CAP-057 | Resolución de red Win32. | CAP-054 | Puerto de infraestructura y health/readiness. | RETIRAR |
| EXP-CAP-058 | Gestión de procesos Win32. | CAP-055 | Jobs/servicios supervisados y observables. | RETIRAR |
| EXP-CAP-059 | Drag & drop OLE y controles ActiveX. | CAP-019 | Interacción web accesible que conserva la jerarquía. | RETIRAR |
| EXP-CAP-060 | Globals y singletons mutables de sesión/entorno. | CAP-037, CAP-053 | Estado por request/sesión y dependencias explícitas. | RETIRAR |
| EXP-CAP-061 | Cambio interactivo de backend y rutas embebidas. | CAP-031 | Configuración tipada por despliegue. | RETIRAR |
| EXP-CAP-062 | Popup modal anti-spam/busy. | CAP-035, CAP-036 | Idempotency keys, controles deshabilitados y `aria-busy`. | RETIRAR |
| EXP-CAP-063 | Menú JSON-hub y navegación acoplada a formularios. | inventario de navegación | Rutas SSR autorizadas por capability. | RETIRAR |

## Dependencias y orden de especificación

1. `EXP-CAP-043..050` fijan identidad, autorización, configuración y readiness; dependen del contrato de Lanzadera.
2. `EXP-CAP-001..024` fijan agregado, relacionados y catálogos sobre PostgreSQL con schema propio.
3. `EXP-CAP-025..032` proyectan consultas y escritura resiliente sin acoplar UI al dominio.
4. `EXP-CAP-033..042` y `EXP-CAP-051..056` cierran contratos externos, hash, batch y trazabilidad.
5. `EXP-CAP-057..063` solo se retiran cuando su sustitución tenga prueba y aceptación.

## Obligaciones para las siguientes fases

- Cada spec citará sus `EXP-CAP-*`, reglas, permisos, estados vacíos, errores e idempotencia; ninguna volverá al código Access para descubrir alcance.
- El diseño deberá completar topología, modelo agnóstico, mapping campo a campo, reconciliación, rechazos, rollback y observabilidad.
- Las tasks deberán separar slices autónomos de ≤400 líneas; las cadenas se justificarán únicamente por dependencia real.
- UAT y cutover probarán paridad de capacidades, no semejanza visual. Toda retirada exige evidencia de sustitución o aprobación explícita.
- No se copiarán PII, rutas locales, secretos ni detalles de formularios al modelo objetivo o a fixtures.
