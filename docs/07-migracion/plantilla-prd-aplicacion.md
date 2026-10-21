# Plantilla del PRD de una aplicación

> **Roadmap:** [roadmap-descubrimiento-prd.md](roadmap-descubrimiento-prd.md) · **Issue:** #1069

Esta plantilla fija la estructura de `docs/03-aplicaciones/<app>/prd/`. Todos los estudios la siguen para que los ocho PRD sean comparables y para que la migración pueda planificarse sin releer el legacy.

## Reglas de redacción

- Cada afirmación sobre el comportamiento cita `ruta:línea` del commit fijado. Sin cita, es pendiente.
- Se describe lo que hace el código, no lo que debería hacer. Las sospechas van a **Posibles defectos**; las mejoras, a **Propuestas web pendientes de decisión**.
- Las referencias dinámicas no resueltas (SQL concatenado, nombres en variables, `Eval`, menús en JSON) se listan como pendientes, con la línea donde aparecen.
- No se citan datos de negocio, credenciales ni rutas de red. Los ejemplos usan valores ficticios.
- Castellano peninsular formal; identificadores del código tal como aparecen en `src`.

## `prd/README.md`

```markdown
# PRD — <Aplicación>

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/<repo> |
| Commit fijado | <SHA completo> (<rama o release>, <fecha>) |
| Índice CodeGraph-VBA | <ficheros VBA indexados>, versión <x.y.z> |
| Estudio previo | <enlaces a epic.md y walkthroughs> |
| Issue | #<n> |

## Inventario reconciliado

| Objeto | Tipo | Ruta | Estado | Ficha o motivo |
|---|---|---|---|---|
| Form_X | formulario | src/forms/Form_X.cls | analizado / no aplicable / pendiente | formularios/Form_X.md |

Totales: <n> analizados · <n> no aplicables · <n> pendientes.

## Delta frente al estudio previo

| Capacidad o formulario | Clasificación | Evidencia |
|---|---|---|
| … | añadida / modificada / retirada / sin cambios | ruta:línea |

## Capacidades

| ID | Capacidad | Formularios | Criterios | Issue propuesta |
|---|---|---|---|---|

## Discrepancias abiertas

Contradicciones entre `src` y la documentación previa, y referencias dinámicas sin resolver.
```

## `prd/formularios/<Formulario>.md`

Una ficha por formulario. Si un grupo de formularios es idéntico salvo el origen de datos, una ficha puede cubrir el grupo siempre que tenga una fila por formulario en la tabla de controles.

```markdown
# <Formulario>

- **Fuente:** src/forms/<Formulario>.cls · src/forms/<Formulario>.form.txt
- **Propósito:** una frase, con cita.
- **Origen de datos:** RecordSource / consultas / tablas, con cita.
- **Se abre desde:** formularios, menús o macros que lo abren, con cita.
- **Abre:** formularios, informes o procesos que lanza, con cita.
- **Permisos:** quién puede abrirlo y qué controles cambian según el rol.

## Controles y eventos

| Control | Tipo | Evento | Handler | Llamadas | Datos y efectos | Capacidad |
|---|---|---|---|---|---|---|
| cmdGuardar | CommandButton | Click | Form_X.cls:120 | modY.Guardar (modY.bas:40) | UPDATE TbZ (modY.bas:55) | CAP-<n> |

Los controles sin evento se listan si tienen origen de datos, validación o valor por defecto.

## Flujos

- **Normal:** pasos numerados, con cita en cada paso.
- **Alternativos:** variantes según datos, rol o estado.
- **Errores:** qué valida, qué mensaje muestra y qué deja en los datos cuando falla.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo (DAO, ADO, SQL guardada, SQL dinámico) | Transacción | Evidencia |
|---|---|---|---|---|

## Coexistencia

Qué debe adaptarse para PostgreSQL vinculado (sintaxis Jet/ACE, funciones, tipos, autonumérico, transacciones) y qué comportamiento debe preservarse. Enlaza a `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

Tres listas separadas. Una propuesta no es requisito hasta su aprobación.

## Criterios de aceptación

Criterios verificables de la versión web, cada uno trazado a una fila de la tabla de controles o a una regla.

## Pendientes

Eslabones de la trazabilidad sin evidencia, con lo que falta para resolverlos.
```

## `prd/erd-origen.md`

Por cada backend: tablas, columnas, tipo, longitud, precisión y escala, nulabilidad, valor por defecto, clave primaria y candidatas, índices, relaciones con reglas de actualización y borrado, y los formularios y módulos que leen o escriben cada tabla. Las relaciones se marcan como **declaradas** (metadatos o DDL) o **inferidas por código** (con cita). Las carencias de metadatos se listan con la extracción de solo lectura que las resolvería.

## `prd/matriz-campos.md`

| Backend | Tabla | Columna | Tipo origen | Destino (esquema.tabla.columna) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|

El destino es una propuesta hasta que se apruebe el ERD global. **Estado**: con destino o pendiente con motivo.

## `prd/coexistencia.md`

Resumen por aplicación de los apartados de coexistencia del roadmap: vínculos y conexiones, adaptaciones de SQL y transacciones, reglas compartidas, concurrencia y política de conflictos, seguridad y auditoría, pruebas bidireccionales y plan de corte con reversión.

## `prd/issues-propuestas.md`

| Propuesta | Capacidad | Evidencia | Depende de | Criterios de aceptación | Etiquetas |
|---|---|---|---|---|---|

Las propuestas no se publican como issues hasta que el PRD esté cerrado. No se duplica una capacidad que ya cubra otra aplicación; se enlaza.
