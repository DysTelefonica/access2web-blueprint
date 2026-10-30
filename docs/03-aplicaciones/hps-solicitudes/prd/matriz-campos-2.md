# Matriz campo a campo (2/2) — HPS Solicitudes

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Issue:** #1077 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md) · **Continuación de [matriz-campos.md](matriz-campos.md)**

Las 5 tablas locales restantes y el anexo de las 9 tablas linked, con la misma convención que la primera entrega. Ninguna fila está aprobada: el ERD de origen queda pendiente del propietario (D-HPS-1 del README), porque `dysflow` no está disponible en el entorno de trabajo.

## Tablas cubiertas en esta entrega (5)

- [Copia de TbExpedientes](#copia-de-tbexpedientes)
- [TbHPSGrado](#tbhpsgrado)
- [TbLogs](#tblogs)
- [TbLogsGeneral](#tblogsgeneral)
- [TbUltimoCambio](#tbultimocambio)

## Anexo — 9 tablas linked (`LINKED/SKIPPED` en el volcado)

Las 9 tablas siguientes no se incluyeron en `ERD/schema.sql` por estar enlazadas a ficheros `.accdb` ajenos que no estaban en el snapshot; aquí se declaran como `[link — pendiente]` con la evidencia disponible en `src`. D-HPS-1: pendiente del workflow `dysflow-erd` (#1144) en la máquina Windows del propietario.

- [TbExpedientes](#tbexpedientes-link)
- [TbExpedientesLugaresEjecucion](#tbexpedienteslugares-ejecucion-link)
- [TbHPS](#tbhps-link)
- [TbHPSEquivalencia](#tbpsequivalencia-link)
- [TbHPSGrado1](#tbhpsgrado1-link)
- [TbMotivoHPS](#tbmotivohps-link)
- [TbUsuarios](#tbusuarios-link)
- [TbUsuariosEntidades](#tbusuariosentidades-link)
- [TbUsuariosHistoricos](#tbusuarioshistoricos-link)

## Reglas de mapeo de tipos

Mismas que en la primera entrega ([matriz-campos.md § Reglas de mapeo de tipos](matriz-campos.md#reglas-de-mapeo-de-tipos-propuesta)). Resumen: `LONG` → `INTEGER`, `INT` → `SMALLINT`, `TEXT(n)` → `VARCHAR(n)`, `TEXT(2)` → `BOOLEAN` con CHECK sobre el dominio (`Sí`/`No`), `MEMO` → `TEXT`, `BOOLEAN` → `BOOLEAN`, `SHORT_DATE_TIME` → `TIMESTAMP` con `WITH TIME ZONE`.

---

<a id="copia-de-tbexpedientes"></a>

### Copia de TbExpedientes

`Rows: 70`. Tabla con sufijo `Copia de` (sentinel de esquema Access: indica copia previa a un cambio masivo, ver `coexistencia.md`). 9 columnas, todas relacionadas con la tabla linked `TbExpedientes` (no local). Mutada por la capa de read-only de `Expediente.cls` y `ExpedientesBusqueda.bas`; el traspaso ONS actualiza `URLAdjuntoEnvioONS` en `TbSolicitudes` y deja este snapshot sin tocar.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDExpediente` | `LONG(4)` | `INTEGER` (migrar como vista o ignorar) | `Form_FormExpedienteDetalle.cls:54-69`; `Form_FormExpedientesBusqueda.cls:233-238` |
| `Cod_Exp` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormExpedientesBusqueda.cls:233-238` |
| `OC` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormExpedienteDetalle.cls:54-69` |
| `Grado` | `TEXT(255)` | `VARCHAR(255)` | id. |
| `Expediente` | `TEXT(255)` | `VARCHAR(255)` | id. |
| `F_Inicio` | `SHORT_DATE_TIME` | `DATE` | id. |
| `F_Fin` | `SHORT_DATE_TIME` | `DATE` | id. |
| `IDJuridicaContrato` | `LONG(4)` | `INTEGER` | id. |
| `IDExpedienteNuevo` | `LONG(4)` | `INTEGER` | id. |

Migración: descartar como tabla (es sentinel de copia) y leer los datos a través de la vista equivalente sobre la `TbExpedientes` linked. D-HPS-5: validar que ningún consumidor lee esta tabla y no la linked.

---

<a id="tbhpsgrado"></a>

### TbHPSGrado

`Rows: 13`. Catálogo de grados HPS. Dos columnas de `TEXT(255)` que forman una clave compuesta implícita. Mutada por `Form_FormMotivoHPS`, `Form_FormMotivosHPS` y `Form_FormAltaDatosSolicitante` (alta y catálogo).

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `TipoHPS` | `TEXT(255)` | `VARCHAR(255)` (parte de la PK compuesta) | `Form_FormMotivoHPS.cls:22-26`; `Form_FormMotivosHPS.cls:22-26` |
| `Grado` | `TEXT(255)` | `VARCHAR(255)` (parte de la PK compuesta) | id. |

Migración: `GET /hps_grados?tipo=...`; índice en `(TipoHPS, Grado)`. Sentinela `LINKED/SKIPPED` para la linked `TbHPSGrado1` (mismo nombre con sufijo, fuente distinta — no incluir en la web).

---

<a id="tblogs"></a>

### TbLogs

`Rows: 0`. Tabla de log específica de solicitudes, **vacía en el snapshot**. Reservada para diagnóstico de operaciones sobre `TbSolicitudes`; no consumida por `src/`. Mutable por `LogGeneralOperaciones.cls` o por código heredado.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDLog` | `LONG(4)` | `INTEGER PK` | (no en `src/`) |
| `IDSolicitud` | `LONG(4)` | `INTEGER FK` | (no en `src/`; herencia de `TbSolicitudes` por FK implícita) |
| `TipoOperacion` | `TEXT(255)` | `VARCHAR(255)` | (no en `src/`) |
| `Fecha` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL DEFAULT NOW()` | (no en `src/`) |
| `Usuario` | `TEXT(255)` | `VARCHAR(255)` | (no en `src/`) |
| `Descripcion` | `MEMO` | `TEXT` | (no en `src/`) |

Migración: tabla huérfana. Decisión abierta D-HPS-6: eliminarla o consolidarla en `TbLogsGeneral`.

---

<a id="tblogsgeneral"></a>

### TbLogsGeneral

`Rows: 2058`. Log general con `DNI` y `Usuario` literales — datos personales en plano. Consumida por `LogGeneral.cls`, `LogGeneralOperaciones.cls` y por `ComandoInformeControl_Click` de `Form_FormCorreosGestion` para registrar visualizaciones de informes.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `IDLog` | `LONG(4)` | `INTEGER PK` | `classes/LogGeneral.cls:1-50`; `RegistrarLogGeneral` (helper) |
| `Fecha` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL DEFAULT NOW()` | `RegistrarLogGeneral` |
| `Usuario` | `TEXT(255)` | `VARCHAR(255)` (FK a `users.id`) | `Form_FormCorreosGestion.cls:42` (`m_ObjUsuarioConectado.UsuarioRed`) |
| `Accion` | `TEXT(255)` | `VARCHAR(255)` | `Form_FormCorreosGestion.cls:42` |
| `Descripcion` | `MEMO` | `TEXT` | id. |
| `DNI` | `TEXT(255)` | `VARCHAR(255)` (seudonimizar) | (escrito por la capa de log al recibir DNI en algún handler — no en `src/`) |
| `CorreosAutomaticos` | `TEXT(2)` | `BOOLEAN` | (no en `src/`; heredado) |
| `IDSolicitud` | `LONG(4)` | `INTEGER FK` | (no en `src/`) |

Migración: partición por `Fecha` (mensual/semestral) en PostgreSQL para los 2058+ filas; índice en `(IDSolicitud)` y `(Usuario, Fecha)`. `DNI` y `Usuario` seudonimizados o hasheados (D-HPS-3).

---

<a id="tbultimocambio"></a>

### TbUltimoCambio

`Rows: 5`. Log minimal de últimos cambios por solicitud, con FK implícita a `TbSolicitudes`. Mutada por `UltimoCambioOperaciones.cls`.

| Columna | Tipo origen | Tipo objetivo | Evidencia en `src` |
|---|---|---|---|
| `ID` | `LONG(4)` | `INTEGER PK` | `classes/UltimoCambioOperaciones.cls:1-50` |
| `IDSolicitud` | `LONG(4)` | `INTEGER FK` | id. |
| `FechaCambio` | `SHORT_DATE_TIME` | `TIMESTAMP NOT NULL` | id. |
| `IDUsuarioCambio` | `LONG(4)` | `INTEGER FK` (a linked `TbUsuarios`) | id. |

Migración: tabla pequeña (≤5 filas en el snapshot); se conserva tal cual. Decisión D-HPS-7: validar que la FK numérica a `TbUsuarios` linked se reemplaza por FK al `users.id` de la web.

---

<a id="tbexpedientes-link"></a>

### TbExpedientes (linked)

Tabla enlazada a `Expedientes_datos.accdb` (externa, no presente en el snapshot). Contiene la versión canónica de los expedientes; la local `Copia de TbExpedientes` es sentinel de copia. Mutada y consultada por la capa de Lanzadera (no presente en `src/` de HPS Solicitudes).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| `IDExpediente` | `LONG(4)` PK | `Form_FormExpedienteDetalle.cls:54-69`; `Form_FormExpedientesBusqueda.cls:233-238` (referenciada por `TbSolicitudes.IDExpediente`) |
| (resto) | (no volcada) | (no en `src/`) |

D-HPS-1 + D-HPS-5: la fuente `Expedientes_datos.accdb` debe extraerse desde Windows con `dysflow`; mientras tanto, lectura a través de `constructor.getExpedientesBusqueda` indirecta.

---

<a id="tbexpedienteslugares-ejecucion-link"></a>

### TbExpedientesLugaresEjecucion (linked)

Lugares de ejecución de un expediente, enlazada a `Expedientes_datos.accdb`. 0 referencias en `src/` de HPS Solicitudes: la pantalla que la usaba probablemente vive en Lanzadera (no en este `.cls`/`.form.txt`).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | (no en `src/`) |

D-HPS-1: pendiente extracción.

---

<a id="tbhps-link"></a>

### TbHPS (linked)

Tabla de datos de HPS propiamente dicha (no el catálogo de grados), enlazada a `HPST.accdb`. Consumida por la web de HPS (no por VBA de HPS Solicitudes) — la referencia aquí es a través de `TbSolicitudes.IDUsuarioHPS` (FK implícita).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | `Form_FormSolicitudRenovacionDatos.cls:640-643` (`IDUsuarioHPS`) |

D-HPS-1: pendiente; la web de HPS es fuente de verdad y la integración es vía API, no vía réplica.

---

<a id="tbpsequivalencia-link"></a>

### TbHPSEquivalencia (linked)

Equivalencias entre identificadores de HPS Solicitudes y los de la web de HPS. Consumida por la capa de integración (no en `src/` de HPS Solicitudes).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | (no en `src/`) |

D-HPS-1: pendiente; al ser una tabla de equivalencias, su rol en la web se cubre con un endpoint `GET /equivalencias?legacy_id=...`.

---

<a id="tbhpsgrado1-link"></a>

### TbHPSGrado1 (linked)

Variante del catálogo `TbHPSGrado` (mismo nombre con sufijo `1`, fuente distinta). 0 referencias en `src/`; sentinel de variante de esquema.

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | (no en `src/`) |

D-HPS-1: confirmar si la linked está duplicada con la local; probablemente descartable.

---

<a id="tbmotivohps-link"></a>

### TbMotivoHPS (linked)

Catálogo de motivos HPS enlazado a `HPST.accdb`. Consumida por `Form_FormMotivoHPS` y `Form_FormMotivosHPS` (alta y catálogo).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | `Form_FormMotivoHPS.cls:22-26`; `Form_FormMotivosHPS.cls:22-26`; `MotivoHPS.EditarMotivo` (UPDATE en `TbSolicitudes`, `TbUsuarios`, `TbUsuariosHistoricos`) |

D-HPS-1: pendiente; los formularios de motivos hacen referencia al nombre lógico `m_ObjEntorno.ColTiposCorreos` que no es exactamente este catálogo. A verificar si la lectura del catálogo de motivos viene de esta tabla o de la columna `Motivo_HPS` de `TbSolicitudes` (que es local).

---

<a id="tbusuarios-link"></a>

### TbUsuarios (linked)

Identidad de los usuarios del dominio, enlazada a `HPST.accdb`. Consumida por `Form_FormAltaDatosSolicitante` (en el DNI/NIE sintético) y por `Form_FormMotivoHPS` (UPDATE en `TbUsuarios`).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | `Form_FormAltaDatosSolicitante.cls:218-220, 292-296`; `MotivoHPS.cls:41-44, 180-182, 208-213` (UPDATE) |

D-HPS-1: pendiente; la web de Lanzadera gestiona la identidad real (`TbUsuariosAplicaciones`), y la de HPS usa su propio `TbUsuarios`. La integración de identidad se hace vía API o, en la web consolidada, vía `users.id` único (decisión abierta, D-HPS-8).

---

<a id="tbusuariosentidades-link"></a>

### TbUsuariosEntidades (linked)

Relación `N-M` entre usuarios y entidades, enlazada a `HPST.accdb`. 0 referencias en `src/`; la gestión de entidades se hace por `Form_FormResponsablesGestion` y `Form_FormResponsable` que apuntan a `TbResponsables` (local).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | (no en `src/`) |

D-HPS-1: pendiente; al ser `N-M`, en la web se traduce a una tabla `user_entidades` con FKs a `users.id` y `entidades.id`.

---

<a id="tbusuarioshistoricos-link"></a>

### TbUsuariosHistoricos (linked)

Histórico de usuarios, enlazada a `HPST.accdb`. Mutada por `MotivoHPS.EditarMotivo` (UPDATE conjunto al renombrar motivos).

| Columna | Tipo origen | Evidencia en `src` |
|---|---|---|
| (volcado pendiente) | — | `MotivoHPS.cls:208-213` (UPDATE) |

D-HPS-1: pendiente; en la web se consolida en una tabla `users_historico` con FK a `users.id` y `modified_by_user_id`.

## Pendientes

- D-HPS-1 (global): extracción de solo lectura del backend `Solicitudes_HPS_datos.accdb` y de los `.accdb` linked con `dysflow-erd` en Windows para confirmar claves primarias, índices, nulabilidad y FKs de las 16 tablas (11 locales + 5 linked que sí tienen contenido en `src/`).
- D-HPS-5: validar que ningún consumidor de `Copia de TbExpedientes` lee esa local y no la linked.
- D-HPS-6: decidir destino de `TbLogs` (vacía) — eliminar o consolidar en `TbLogsGeneral`.
- D-HPS-7: validar FK numérica a `TbUsuarios` linked desde `TbUltimoCambio.IDUsuarioCambio` y decidir mapeo a `users.id` web.
- D-HPS-8: consolidar la identidad de las linked `TbUsuarios`/`TbUsuariosEntidades`/`TbUsuariosHistoricos` con la de Lanzadera (decisión abierta del bloque P-EXP-3 de la coexistencia global).
- Mapeo de `TipoHPS`/`Grado` y de las linked `TbHPSGrado1`/`TbMotivoHPS` (decidir si la web sirve un solo catálogo o dos).
- Política de retención de `TbLogsGeneral` (2058 filas en el snapshot): partición por `Fecha` o archivo en frío.
- Decisión de sí/no en las linked `TbHPS` (no leídas en `src/` pero referenciadas por `TbSolicitudes.IDUsuarioHPS`): ¿se migran como linked, se consolidan o se eliminan?
- Verificación de FKs físicas: todas las del inventario son FKs «implícitas» inferidas por la convención de nombre de columna + uso en `src/`; la validación formal requiere `dysflow`.
