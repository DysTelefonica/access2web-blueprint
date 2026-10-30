# Issues propuestas — Gestión de Riesgos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · **Fase:** 6 · **PRD:** [README](README.md)

Unidades de trabajo que este PRD propone. **No se ha creado ninguna de ellas.** Cada una es un problema con dueño, criterio de cierre y evidencia; ninguna duplica una capacidad de otra aplicación del roadmap.

Las que dependen del propietario o de otra aplicación se marcan y se quedan abiertas a su decisión.

## Resumen

| Id | Propuesta | Tipo | Depende de |
|---|---|---|---|
| P-GR-1 | Extraer el ERD de origen del backend | propietario | Mergear #1144 + ruta del `.accdb` |
| P-GR-2 | Revisar las credenciales en `Usuario.cls` | propietario, **no publicar** | — |
| P-GR-3 | Cerrar el vínculo a sí mismo de `TbProyectos1` | datos | P-GR-1 |
| P-GR-4 | Decidir el modelo del árbol de riesgos en la web | decisión | — |
| P-GR-5 | Decidir si las tablas de caché se migran | decisión | P-GR-4 |
| P-GR-6 | Fijar la estrategia de identificadores | datos | P-GR-1 |
| P-GR-7 | Extraer el contrato de lectura que la web debe servir al legacy | web | — |
| P-GR-8 | Sustituir el HTML en disco por una ruta servida y autenticada | web | — |
| P-GR-9 | Caracterizar el comportamiento de error real en UAT | UAT | fichas |
| P-GR-10 | Definir el corte y la reversión | decisión | P-GR-1, P-GR-7 |
| P-GR-11 | Decidir el destino del trabajo sin liberar de staging | decisión | — |
| P-GR-12 | Retirar objetos huérfanos del legacy | legacy | — |

## P-GR-1 · Extraer el ERD de origen del backend

**Qué:** extraer, en solo lectura, las tablas, claves primarias, índices, relaciones y nulabilidad de `Gestion_Riesgos_Datos.accdb`.

**Por qué:** sin esto, las 85 filas de la [matriz](matriz-campos.md) son propuesta y las relaciones son inferidas por código. Es el bloqueo raíz de la fase 3.

**Quién:** el propietario, en su máquina Windows, con `dysflow` y el backend accesible al runner.

**Bloqueos, medidos hoy:** el backend **no está en el repositorio** (el clon solo trae el frontend `Gestion_Riesgos.accdb`, y `mdb-tables` lo rechaza); y el workflow de extracción está en el **PR #1144, abierto y sin mergear**. Ya anotado en la [#1079](https://github.com/DysTelefonica/access2web-blueprint/issues/1079).

**Cierre:** `prd/erd-origen.md` con la extracción, las filas de la matriz pasadas de «propuesta» a «confirmada», y D-GR-4 resuelta. Debe coordinarse con las otras siete aplicaciones por las tres tablas compartidas de Lanzadera.

## P-GR-2 · Revisar las credenciales de `src/classes/Usuario.cls`

**Qué:** el fichero contiene literales de tipo contraseña o clave. No se reproducen aquí y **esta issue no se publica en el repositorio del legacy**.

**Por qué:** es un hallazgo de esta revisión documental, no un defecto de negocio. La decisión de rotar, retirar o reubicar es del propietario.

**Cierre:** el propietario confirma el alcance y decide la acción.

## P-GR-3 · Resolver `TbProyectos1` declarado como vinculado a sí mismo

**Qué:** `ERD/Estructura_Datos.md` declara `TbProyectos1` como tabla **vinculada** a `Gestion_Riesgos_Datos.accdb`, que es el propio backend de la aplicación.

**Por qué:** un vínculo a sí mismo no puede ser. O el documento está desactualizado (es de 2026-02-13 y el commit fijado es de julio), o hay algo que este estudio no entiende. Mientras no se aclare, esa tabla no sirve como evidencia de vínculo.

**Cierre:** confirmado contra el binario; D-GR-5 resuelta.

## P-GR-4 · Decidir el modelo del árbol de riesgos en la web

**Qué:** hoy el árbol se materializa en seis tablas de caché y se reconstruye a mano en cada cambio de edición, riesgo, plan o materialización. El estudio previo ya señalaba la lentitud (D88).

**Por qué:** en PostgreSQL el árbol debería salir de una CTE recursiva sobre los datos. La alternativa —tabla de presentación mantenida por disparador— tiene su propio coste de escritura.

**Decisión del propietario.** Este PRD **no recomienda** entre las dos: depende de la carga real, que no se ha medido.

**Cierre:** modelo elegido y documentado con su coste esperado.

## P-GR-5 · Decidir si las tablas de caché se migran

**Qué:** `TbArbolCacheMeta`, `TbCacheArbolRiesgosMeta`, `TbCacheArbolRiesgosNodo`, `TbCacheControlCambiosMeta`, `TbCacheControlCambiosRow` y `TbCachePublicabilidadEdicion`.

**Por qué:** están en tensión directa con P-GR-4: migrarlas perpetuaría el problema que las motivó.

**Cierre:** cada tabla con destino decidido: migrar, derivar o descartar.

## P-GR-6 · Fijar la estrategia de identificadores

**Qué:** `DameID` calcula el siguiente identificador como **máximo más uno** sobre el nombre de tabla que recibe por cadena (`src/modules/Funciones Generales.bas:2191, 2223`).

**Por qué:** con dos usuarios guardando a la vez, ambos pueden leer el mismo máximo y escribir el mismo identificador.

**Bloqueo:** si el ERD de origen confirma una restricción de unicidad, el conflicto se rechaza y el riesgo es menor de lo que parece. **La conclusión depende de P-GR-1** y así se declara.

**Cierre:** estrategia de identificadores decidida y documentada (secuencia, identidad o clave natural), con su política de migración.

## P-GR-7 · Extraer el contrato de lectura que la web debe servir al legacy

**Qué:** enumerar el conjunto real de consultas que esta aplicación ejecuta contra cada tabla **vinculada** —15 tablas de seis backends— mientras conviven, y comprobarlo contra la API de la web.

**Por qué:** no hay consultas guardadas: **todo el SQL está embebido** en 195 ficheros de producción. No hay un punto único donde revisar qué necesita el legacy. Es trabajo, no un trámite, y bloquea el corte.

**Cierre:** inventario de consultas con su tabla destino y su prueba de equivalencia.

## P-GR-8 · Sustituir el HTML en disco por una ruta servida y autenticada

**Qué:** hoy los informes se generan como fichero y se abren con el visor del sistema (`Form_FormInformeTipoSalida.cls:81, 85`) o con un control de navegador embebido que lee una ruta local (`Form_FormWeb.cls:32, 38`). El visor además **cierra en silencio** si el fichero no existe (H-GR-9).

**Por qué:** acopla el despliegue de la web al sistema de ficheros del usuario y salta cualquier control de acceso.

**Cierre:** ruta servida y autenticada decided, con el comportamiento de «recurso no encontrado» definido.

## P-GR-9 · Caracterizar el comportamiento de error real en UAT

**Qué:** la aplicación tiene **263** usos de `On Error Resume Next` frente a **16** de `On Error GoTo 0`, y **3 752** `Err.Raise` que propagan por cadena.

**Por qué:** el comportamiento de error es observable y buena parte es indeducible del código. **Sin caracterizarlo no se pueden escribir criterios de aceptación web que signifiquen algo**: no se sabe qué hace hoy el sistema ante un fallo de base de datos, una ruta perdida o un conflicto.

**Cierre:** una matriz de escenarios de fallo con el comportamiento real observado, convertida en criterio de aceptación.

## P-GR-10 · Definir el corte y la reversión

**Qué:** el corte de esta aplicación **no puede ser de esta aplicación**. Depende de que Lanzadera tenga identidad y permisos migrados, de que los cuatro backends externos sigan respondiendo con el mismo contrato, y de que el correo siga saliendo (por base de datos, no por SMTP).

**Por qué:** la reversión no está definida y no puede definirse con lo que hay hoy. Un corte deja en producción un legacy que depende de siete backends que pueden haber cambiado.

**Decisión del propietario**, coordinada con P-EXP-3 (arquitectura de datos de coexistencia), que sigue abierta y escalada.

**Cierre:** secuencia, puerta de entrada y reversión documentadas.

## P-GR-11 · Decidir el destino del trabajo sin liberar de staging

**Qué:** el snapshot de staging tiene 140 objetos que `main` no tiene: 112 módulos de prueba y soporte, y **28 objetos de aplicación** (servicios de publicabilidad, detectores de cambios, repositorio de priorización, entre otros). Además hay 3 objetos en `main` que no están en staging.

**Por qué:** varios de los ausentes en `main` (`PriorizacionRepository`, `RiesgoMaterializacionContingenciaService`, `MigracionMaterializacionPlanContingencia`, `MaterializacionPlanEligibilityService`) son justamente la materialización de contingencia que `Form_FormMaterializacionPlanContingencia` daba por cubierta. **Si se liberan, hay que reabrir el análisis de los formularios de plan.**

**Cierre:** cada objeto con decisión de liberar, retirar o mantener en desarrollo.

## P-GR-12 · Retirar objetos huérfanos del legacy

**Qué:** `Form_Formulario1` está en `main`, no lo abre nada y sus controles pertenecen a otra aplicación (H-GR-15 a H-GR-17). `Form_FormMaterializacionPlanContingencia` desapareció de `main` y su cobertura parece haberse repartido entre `Form_FormPlanPrincipal` y `Form_FormPlanAcciones`.

**Por qué:** son objetos que no migran; retirarlos de forma explícita evita que entren en alcance por omisión.

**Decisión del propietario** sobre su repositorio legacy. Este PRD no modifica ese repositorio.

**Cierre:** objetos retirados, o justificación documentada de por qué se conservan.

## Relación con otras aplicaciones

Ninguna de estas propuestas duplica una capacidad de otro aplicativo:

- **Lanzadera** es propietaria de la identidad y los permisos; P-GR-10 y la parte de identidad de P-GR-7 se coordinan con su PRD, no se redefinen aquí.
- **Expedientes** y **No Conformidades** son proveedores de tablas y backends; su corte condiciona el de esta.
- **La arquitectura de datos de coexistencia (P-EXP-3)** sigue siendo una decisión abierta del propietario, en estudio, y no se recomienda en este documento.