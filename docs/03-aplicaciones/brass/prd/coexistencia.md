# Coexistencia — Brass

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Issue:** #1078 · **Fase:** 4 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Qué debe resolverse para que Brass funcione en Access y en la web a la vez sobre los mismos datos. Es medición estática de `src` (sin ejecutar Access); los recuentos son de líneas de código de producción, sin los módulos de prueba (`Test*`, `Pruebas*`, `modTestingSLA`). `VG.bas` abrevia `modules/Variables Globales.bas`; `FG.bas`, `modules/Funciones Generales.bas`. Ninguna propuesta es requisito hasta su aprobación; la arquitectura de datos de coexistencia sigue abierta y escalada al propietario (P-EXP-3 en el PRD de Expedientes).

## 1. Vínculos y conexiones

- **Doble mecanismo de acceso a datos.** Brass usa tablas vinculadas del propio `.accdb` (226 líneas con `CurrentDb`, p. ej. `Constructor.bas:1365`) y, a la vez, abre el backend por ruta con DAO (`getdb`, unas 180 líneas de llamada). La ruta del backend se obtiene primero del `Connect` de las tablas vinculadas (`VG.bas:474-500`, usada en `getdb`, `VG.bas:548-`); si no hay vínculo, de la variable temporal `DatosEnLocal` («Sí»: carpeta local; «No»: carpeta remota), que el arranque fija a «No» (`VG.bas:293, 527`).
- **Contraseña literal.** `getdb` y `getdbLanzadera` abren las bases con una contraseña escrita en el código (`VG.bas:539, 573`); otra apertura la repite (`FG.bas:1335`). No se transcribe. Con PostgreSQL vinculado esa credencial desaparece del cliente, pero hasta el corte sigue siendo un secreto compartido en un binario (ver §5).
- **Backend de otra aplicación.** `getdbLanzadera` abre el `.accdb` de Lanzadera por ruta (`VG.bas:520-546`) para leer `TbUsuariosAplicaciones` (`Constructor.bas:44-71`) y `TbUsuariosAplicacionesPermisos` (`Constructor.bas:106-110`) y para registrar aperturas y cierres en `TbAplicacionesAperturas`, con lectura y `UPDATE` (`FG.bas:7623-7628, 7668-7672`). Es el contrato de identidad compartido que el PRD de Lanzadera deja abierto (P-LAN-3).
- **Base de otra aplicación.** `FG.bas:1325-1340` abre un `.accde` ajeno por ruta fija para consultar `TbDocumentos` (integración con AGEDO). Sin equivalente web definido: pendiente de decisión.
- **Formularios.** Ninguno de los 87 `.form.txt` declara `RecordSource` (recuento mecánico): los formularios no están enlazados a tablas, de modo que el acceso a datos está en el código VBA y no en las propiedades del formulario.

## 2. Adaptaciones Jet/ACE y transacciones

| Aspecto | Medida en `src` | Ejemplo | Adaptación a PostgreSQL vinculado |
|---|---|---|---|
| SQL por concatenación | 480 líneas con `SELECT/UPDATE/DELETE` construidos con `&` | `Constructor.bas:1094-1095` | Parametrizar; sin escape de comillas en `Dame` (`FG.bas:290-312`) |
| Literales de fecha `#…#` en formato `mm/dd/yyyy` | 66 usos de `Format(..., "mm/dd/yyyy")` | `Constructor.bas:1094, 1155` | `DATE`/`TIMESTAMP` con parámetro; formato dependiente de la configuración regional |
| `Like '*'` y comodines Jet | 66 usos | `Constructor.bas:2502-2512` | `%`; `LIKE` de Jet no distingue mayúsculas y el de PostgreSQL sí (`ILIKE`) |
| Funciones del motor en SQL | `Date()` en SQL, `DCount/DLookup/DMax` (10) | `Constructor.bas:5133`, `FG.bas:7668` | `CURRENT_DATE`; dominios agregados sustituidos por consultas |
| Identificadores `Max + 1` | `DameID1`, 23 llamadas | `FG.bas:332-363` | Secuencia o identidad en la base; hoy duplicable en concurrencia |
| Altas con `AddNew` | 54 usos; 49 `Execute`; 20 `RunSQL` con 60 `SetWarnings` | `FG.bas:2372, 2970` | Con `SetWarnings` desactivado los errores del motor pueden no mostrarse; revisar cada escritura |
| Transacciones | **0** (`BeginTrans`, `CommitTrans`, `Rollback`) | — | Operaciones de varios pasos (factura con sus tablas de detalle, franqueo) sin atomicidad hoy |
| Manejo de errores | 313 `On Error Resume Next` | — | Errores de PostgreSQL vinculado pueden ocultarse igual que hoy |

## 3. Reglas compartidas

Las reglas viven en VBA, no en el motor: unicidad por consulta previa (p. ej. `Form_FormActividadTipoAlta.cls:96-128`), cálculo de SLA y franqueo en `Evento.cls`, y facturación en `Factura.cls`. Con dos clientes (Access y web) cada regla necesita **una** implementación; la opción por defecto del estudio es mantener la lógica en el servidor web y que Access solo la invoque, pero no está decidida. Hasta que se cierre, toda regla nueva en la web debe replicarse a mano en VBA o bloquear la edición desde Access. Las fichas de formulario y el inventario de módulos y clases listan las reglas con su cita.

## 4. Concurrencia y conflictos

- Sin transacciones ni bloqueo explícito: el patrón dominante es leer, validar y escribir en pasos separados (unicidad comprobada y luego insertada).
- Identificadores `Max + 1` (`FG.bas:350-363`): dos altas simultáneas, una desde Access y otra desde la web, pueden chocar. Exige una única fuente de identificadores (decisión abierta, análoga a P-LAN-6).
- Sellos de modificación escritos con `Now()` del cliente (`Evento.cls:1689`, `Módulo2.bas:428`): con dos clientes, relojes distintos. Pendiente decidir si manda el reloj del servidor.
- Política de conflictos (última escritura gana, versión de fila, bloqueo por registro): por definir.

## 5. Seguridad y auditoría

- **Identidad:** el usuario de red sale de `Wscript.Network` (`Constructor.bas:10-11`, `VG.bas:318-319`) y se busca en `TbUsuariosAplicaciones` de Lanzadera; los permisos salen de `TbUsuariosAplicacionesPermisos` (`Constructor.bas:106-110`). En la web la identidad debe provenir del contrato de Lanzadera, no de este mecanismo.
- **Secretos:** contraseña de acceso a las bases literal en el código (§1); mover a gestión de secretos en el corte. Pendiente confirmar con el propietario si coincide con la de otras aplicaciones (no se comprueba aquí).
- **Auditoría:** solo `FechaRegistroAlta`/`FechaRegistroModificacion` en eventos (`Evento.cls:1689`); no hay tabla de auditoría en el esquema de origen (57 tablas, [ERD/schema.sql](../ERD/schema.sql)). La web debería auditar escrituras de ambos clientes: propuesta pendiente.
- **Datos personales:** alias de técnico, usuario y contacto (clase «personal» de la [matriz](matriz-campos.md)).

## 6. Pruebas bidireccionales

Propuesta, sin aprobar: para cada capacidad que escriba (alta y edición de evento, franqueo con SLA, facturación, calibración), pares de casos «escribo en Access, leo en la web» y a la inversa, sobre una copia de datos y con el mismo conjunto de reglas. Los módulos `Pruebas`, `PruebasSLA` y `Test_*` existentes cubren SLA, no integración. La suite de regresión de VBA que valide estas reglas debe correr contra el `.accdb` de prueba, nunca contra producción.

## 7. Corte en la Lanzadera y reversión

Brass se abre hoy desde la Lanzadera como una aplicación más; qué significa «lanzar» una aplicación que ya está en la web durante la coexistencia lo decide el propietario (P-LAN-5). El corte por capacidades (de lectura a escritura) y la **reversión quedan por definir**: sin ellas no se puede planificar la salida. Condiciones previas: ERD de origen confirmado (D-BRA-1), esquema web de Brass aprobado y fuente única de identificadores.

## Pendientes

- Las rutas de datos de producción y la ubicación del `.accdb` del backend no constan en `src`: las aporta el propietario para la extracción del ERD.
- Tablas vinculadas reales del frontend (`TableDefs`): requieren abrir el binario; no se han leído.
- Verificar en el binario si las contraseñas literales (tres apariciones) coinciden entre sí y si se rotan.
