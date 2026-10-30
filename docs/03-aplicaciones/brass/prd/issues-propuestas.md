# Issues propuestas — Brass

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Issue:** #1078 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Unidades implementables derivadas del PRD. **No se ha creado ninguna issue**: se publican cuando el PRD esté cerrado y con la etiqueta `app/brass`. Ninguna se declara requisito hasta su aprobación. Las capacidades `CAP-n` son las de [capacidades](capacidades.md).

| Propuesta | Capacidad | Evidencia | Depende de | Criterios de aceptación | Etiquetas |
|---|---|---|---|---|---|
| P-BRA-1 Extraer y confirmar el ERD de origen (claves, índices, relaciones, nulabilidad) del backend de Brass | todas | D-BRA-1 del [README](README.md); [matriz](matriz-campos.md) sin claves ni relaciones | workflow `dysflow-erd` (#1144) y ruta del `.accdb` aportada por el propietario | `erd-origen.md` con tablas, columnas, índices y relaciones marcadas declaradas o inferidas; las 419 columnas pasan de «propuesta» a «confirmada» o a pendiente con motivo | `app/brass`, `type:docs` |
| P-BRA-2 Esquema web de Brass y mapeo de tipos | todas | [matriz](matriz-campos.md): 57 tablas sin destino | P-BRA-1; topología PostgreSQL global | Cada tabla y columna tiene destino `esquema.tabla.columna` o disposición justificada; decisión de identificador de evento (`TEXT(50)`) tomada | `app/brass` |
| P-BRA-3 Identificadores sin `Max + 1` | CAP-3, CAP-4, CAP-8 | `Funciones Generales.bas:332-363` (23 llamadas a `DameID1`) | P-BRA-2; decisión de fuente única de identificadores (análoga a P-LAN-6) | Altas concurrentes desde Access y web no repiten identificador; prueba bidireccional | `app/brass` |
| P-BRA-4 Contrato de identidad y permisos compartido con Lanzadera | CAP-1 | `Variables Globales.bas:335-347`, `Constructor.bas:44-110` | P-LAN-3 (no se duplica; se enlaza) | La web obtiene usuario y permisos (administrador o técnico) del contrato de Lanzadera; el acceso denegado reproduce la regla de `EVE` | `app/brass` |
| P-BRA-5 Eliminar la contraseña literal de los accesos a bases | CAP-1, CAP-2 | `Variables Globales.bas:539, 573`, `Funciones Generales.bas:1335` | Decisión del propietario sobre gestión de secretos | Ningún secreto en código ni en la web; rotación documentada; el detalle se comunica al propietario, no se publica | `app/brass`, `security` |
| P-BRA-6 Eventos: alta, edición, copia y gestión en la web | CAP-3 | Fichas de `FormEventoAlta`, `FormEventoEdicion`, `FormEventoCopia`, `FormEventoGestion`; `Evento.cls` | P-BRA-2, P-BRA-3 | Paridad de campos y validaciones de las fichas; escrituras atómicas (hoy 0 transacciones) | `app/brass` |
| P-BRA-7 SLA y franqueo de evento | CAP-3, CAP-10 | `Evento.cls:1255`; módulos `SLAValidator`, `SLAReportService`, `SLAHTMLService` | P-BRA-6 | El franqueo valida SLA y motivos no franqueables con los mismos resultados que Access sobre un conjunto de casos | `app/brass` |
| P-BRA-8 Equipos de medida y calibraciones | CAP-5 | Fichas `FormCalibracion*`, `FormEquipo*`; clases `EquipoMedida*` | P-BRA-2 | Estados y fechas de calibración equivalentes; dominio regulatorio preservado | `app/brass` |
| P-BRA-9 Planificación y partes | CAP-6 | Fichas `FormPlanificacion*`, `FormParteGestion`; sin `Planificacion.cls` | P-BRA-2 | Calendarios, reprogramación y cierre con los mismos resultados; discrepancia con el estudio previo cerrada | `app/brass` |
| P-BRA-10 Facturación, gastos y subcontratación | CAP-8 | Fichas `FormFactura*`, `FormGasto*`, `FormSubContratacion*`; `Factura.cls` | P-BRA-2, P-BRA-3 | Facturas con sus cinco tablas de detalle consistentes; importes con escala definida; 0 transacciones hoy, atómico en la web | `app/brass` |
| P-BRA-11 Materiales, reparaciones y repuestos | CAP-7 | Fichas `FormMaterial*` | P-BRA-2 | Paridad de reparación, repuestos e intervenciones | `app/brass` |
| P-BRA-12 Técnicos, ausencias, fiestas y precios | CAP-9 | Fichas `FormTecnico*`, `FormTipoTecnicoPrecio*` | P-BRA-2 | Horas, ausencias y precios por tipo con paridad; datos personales tratados según la clase «personal» | `app/brass` |
| P-BRA-13 Informes (SLA, RAC, desde lista, por factura) | CAP-10 | Fichas `FormInforme*`; `Informe.cls` | P-BRA-7 | Los informes se generan con los mismos totales; la generación de HTML no depende de rutas locales | `app/brass` |
| P-BRA-14 Anexos, estructura y catálogos | CAP-11, CAP-12 | Fichas `FormAnexos*`, `FormBUIGestion`, `FormNodo*`, `FormSubSistema*` | P-BRA-2 | Anexos accesibles sin ruta de red del cliente; catálogos editables con las mismas restricciones | `app/brass` |
| P-BRA-15 Coexistencia: SQL, concurrencia y pruebas bidireccionales | todas | [coexistencia](coexistencia.md) | P-EXP-3 (arquitectura de datos de coexistencia, abierta) | Literales de fecha, `Like '*'` y SQL concatenado revisados; pares de pruebas Access↔web por capacidad que escribe | `app/brass` |
| P-BRA-16 Corte en la Lanzadera y reversión | todas | [coexistencia](coexistencia.md), apartado 7 | P-LAN-5; P-BRA-1 a P-BRA-15 | Plan de corte por capacidad con reversión probada | `app/brass` |
| P-BRA-17 Integración con la base documental ajena (AGEDO) | sin asignar | `Funciones Generales.bas:1325-1340` | Decisión del propietario | Decisión registrada: retirar, sustituir o mantener la consulta de documentos | `app/brass` |

## No duplicadas con otras aplicaciones

La identidad y los permisos dependen del contrato de Lanzadera (P-LAN-3) y la arquitectura de datos de coexistencia de P-EXP-3: se enlazan, no se rehacen.

## Pendientes

- Priorización y tamaño de cada propuesta: del propietario.
- P-BRA-17: no consta en `src` qué consulta se hace ni su uso real; pendiente de verificar.
