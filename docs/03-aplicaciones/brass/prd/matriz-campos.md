# Matriz campo a campo — Brass

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Issue:** #1078 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de origen del backend de Brass, con la evidencia de su uso en `src`. Ninguna fila está aprobada: la topología PostgreSQL global está abierta y **no existe todavía un esquema web de Brass** al que apuntar (en `app/src/modules` solo hay módulos de Expedientes y Lanzadera). El destino de cada columna queda por tanto pendiente, y el tipo de destino es una propuesta mecánica.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), extracción del backend `Gestion_Brass_Gestion_Datos.accdb` del snapshot (57 tablas locales, 419 columnas, 0 tablas vinculadas). **No está revalidada contra el backend del commit fijado**: el ERD de origen queda pendiente del propietario (D-BRA-1 del README).
- **Qué falta en esa fuente:** solo columna y tipo. No hay claves primarias, índices, relaciones, nulabilidad ni valores por defecto. El [estudio previo](../data-model.md) afirma 27 relaciones y da claves y obligatoriedad de algunas tablas, pero sin reproducibilidad: no se toma como hecho. Esas columnas se completarán con la extracción de solo lectura.
- **Evidencia en `src`:** se busca cada columna en los ficheros de producción (formularios, módulos y clases, excluidas las pruebas) que citan la tabla. Es lectura estática; los nombres de tabla construidos en variables no se resuelven.
- **Sin valores:** la matriz no contiene datos de ninguna tabla, solo estructura y recuentos de filas del snapshot.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `INTEGER` | Identificadores de otras tablas son `TEXT(50)` en eventos: ver decisión de identidad abajo |
| `INT` | `SMALLINT` | — |
| `TEXT(n)` | `VARCHAR(n)` | `Sí`/`No` en texto de longitud 2: conservar hasta decidir `BOOLEAN` |
| `MEMO` | `TEXT` | — |
| `DOUBLE`, `FLOAT` | `DOUBLE PRECISION` | Importes en facturación y gastos: decidir `NUMERIC` con escala |
| `SHORT_DATE_TIME` | `TIMESTAMP` | Fecha o fecha y hora; zona horaria |
| `BOOLEAN` | `BOOLEAN` | — |

Decisión abierta de identidad: las claves de evento y actividad son texto (`TEXT(50)`); mantener cadena, UUID o numérico depende del ERD global y de la coexistencia.

## Clase de dato

Heurística por nombre de columna, sin revisar contenido: `credencial` (contraseñas, claves), `personal` (alias de técnico, usuario, nombre, correo, teléfono, contacto), `económico` (importes, precios, costes) o «—». Sirve para priorizar la revisión de privacidad y de acceso; no sustituye a una clasificación.

## Leyenda de la columna «Evidencia en `src`»

- `F`: la columna aparece como `Fields("columna")`, `!columna` o `[columna]` en un fichero que cita la tabla (primer hallazgo).
- `S`: el nombre aparece en un fichero que cita la tabla, sin acceso por campo; coincidencia de nombre, poco fiable para nombres comunes.
- `—`: sin cita en los ficheros de producción.

## Tablas de otros backends que Brass lee o escribe

Brass abre además el backend de Lanzadera con `getdbLanzadera` (`Variables Globales.bas:520-546`): lee `TbUsuariosAplicaciones` (`Constructor.bas:44-71`), `TbUsuariosAplicacionesPermisos` (`Constructor.bas:106-110`) y `TbAplicacionesAperturas` (`Funciones Generales.bas:7623-7628`) y actualiza `TbAplicacionesAperturas` al cerrar (`Funciones Generales.bas:7672`). Esas tablas pertenecen a la matriz de Lanzadera (#1074) y no se duplican aquí; el contrato de identidad compartido está pendiente de decisión del propietario.

## Resumen por tabla

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` | Familia | Detalle |
|---|---|---|---|---|---|
| `Tb0FiltroGestion` | 5 | 6551 | 1 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tb0filtrogestion) |
| `TbActividades` | 14 | 25639 | 23 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbactividades) |
| `TbAnexos` | 10 | 7583 | 6 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbanexos) |
| `TbAuxActividad` | 14 | 0 | 1 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbauxactividad) |
| `TbAuxEventos` | 20 | 0 | 1 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbauxeventos) |
| `TbAuxManteniminetosPreventivosCalendario` | 9 | 93 | 0 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbAuxMateriales` | 14 | 0 | 1 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbAuxPlanificacion` | 8 | 81 | 1 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbBUI` | 1 | 18 | 4 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbbui) |
| `TbBuiIDEvento` | 2 | 11 | 1 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbbuiidevento) |
| `TbCausaFin` | 1 | 4 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbcausafin) |
| `TbCodActividad` | 2 | 12 | 7 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbcodactividad) |
| `TbCriticidad` | 2 | 3 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbcriticidad) |
| `TbEquipos` | 7 | 806 | 22 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbEquiposCalibrables` | 7 | 27 | 6 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbEquiposCalibrablesFechas` | 6 | 94 | 5 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbEquiposMedida` | 9 | 14 | 2 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbEquiposMedidaCalibraciones` | 6 | 28 | 2 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbEventos` | 38 | 5773 | 38 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbeventos) |
| `TbEventosEquipoMedida` | 4 | 302 | 4 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbeventosequipomedida) |
| `TbEventosEquipoMedida_antes` | 3 | 26 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbeventosequipomedida_antes) |
| `TbFacturaActividadesInvolucradas` | 18 | 25107 | 7 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaEventosInvolucrados` | 21 | 5514 | 7 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaGastosInvolucrados` | 8 | 519 | 4 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaMaterialesInvolucrados` | 15 | 1986 | 6 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaPrincipal` | 19 | 53 | 4 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaPrincipalPerfiles` | 7 | 304 | 2 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbFacturaSubcontratacionesInvolucradas` | 7 | 74 | 4 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbGastos` | 7 | 519 | 6 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbGastosImportePorTipo` | 6 | 3 | 2 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbGuiaConciliaciones` | 7 | 0 | 0 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbHerramientaDocAyuda` | 2 | 4 | 1 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbherramientadocayuda) |
| `TbMaterial` | 15 | 2351 | 14 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbMaterialSeguimiento` | 11 | 2049 | 6 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbNodoBUI` | 2 | 19 | 6 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbnodobui) |
| `TbNodos` | 1 | 6 | 1 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbnodos) |
| `TbOriginador` | 1 | 7 | 1 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tboriginador) |
| `TbPartesDetalle` | 14 | 8819 | 3 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbPartesPpal` | 8 | 342 | 3 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbPlanificacion` | 13 | 624 | 8 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbPlanificacionAnexos` | 6 | 5 | 2 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbPlanificacionEquipos` | 5 | 82 | 6 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbPlanificacionRegistrada` | 5 | 23 | 2 | Planificación y partes | planificacion-partes (en la tanda siguiente) |
| `TbRepuestosTipo` | 1 | 3 | 0 | Equipos, calibraciones y materiales | equipos-materiales (en la tanda siguiente) |
| `TbResultadoVerificacion` | 1 | 3 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbresultadoverificacion) |
| `TbSistema` | 1 | 2 | 0 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbsistema) |
| `TbSubcontrataciones` | 7 | 75 | 4 | Facturación, gastos y subcontratación | facturacion (en la tanda siguiente) |
| `TbSubsistemaBui` | 2 | 162 | 11 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbsubsistemabui) |
| `TbTecnicos` | 8 | 115 | 6 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbtecnicos) |
| `TbTecnicosAusencias` | 6 | 1236 | 9 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbtecnicosausencias) |
| `TbTecnicosFiestas` | 2 | 104 | 8 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbtecnicosfiestas) |
| `TbTipoAccion` | 1 | 8 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbtipoaccion) |
| `TbTipoAsistencia` | 1 | 4 | 0 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbtipoasistencia) |
| `TbTipoEvento` | 1 | 6 | 1 | Eventos, actividades y anexos | [eventos](matriz-campos/eventos.md#tbtipoevento) |
| `TbTipoTecnico` | 1 | 10 | 4 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbtipotecnico) |
| `TbTipoTecnicoPrecios` | 6 | 19 | 5 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbtipotecnicoprecios) |
| `TbUbicacion` | 1 | 3 | 3 | Técnicos, estructura y configuración | [tecnicos-estructura](matriz-campos/tecnicos-estructura.md#tbubicacion) |

## Tandas de la matriz

La matriz cubre las 57 tablas del backend en cinco ficheros por familia (eventos y actividades; equipos y materiales; planificación y partes; facturación; técnicos y estructura). Lo que depende del ERD de origen (claves, índices, relaciones, nulabilidad, valores por defecto) queda pendiente y se marcará al llegar `erd-origen.md`.
