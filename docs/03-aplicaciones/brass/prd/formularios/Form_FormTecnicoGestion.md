# Form_FormTecnicoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoGestion.cls` (644 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado y gestión de técnicos (`TbTecnicos`). Es un formulario **no enlazado**: la lista `ListaTecnicos` es una lista de valores que `Filtrar` rellena con un `Recordset` DAO (`Form_FormTecnicoGestion.cls:492-578`), filtrando por alias y por «en activo». Desde él se da de alta, edita y elimina un técnico y se navega a sus ausencias, horas, actividades, fiestas y perfiles. Las citas son `Form_FormTecnicoGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todos los controles: CAP-9 (Técnicos, ausencias, fiestas y precios) salvo donde se indica.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpciones` (`lblGestionTecnicos`) | `DoCmd.OpenForm "FormTecnicoGestion"` | `Form_Form0BDOpciones.cls:112-115` |
| `Form0BDInformes` (`lblConsultarHorasTecnico`) | abre este formulario y cierra el de informes | `Form_Form0BDInformes.cls:32-36` |

Salidas: `FormTecnico` (alta y edición, `:22, 99`), `FormActividadGestion` (`:57`), `FormTecnicoAusenciasGestion` (`:274`), `FormTecnicoConsultaHoras` (`:234`), `FormTecnicoFiestasGestion` (`:120`) y `FormTipoTecnicoPreciosGestion` (`:339`). El técnico activo se pasa por la global `m_ObjTecnicoActivo` (`Variables Globales.bas:245`). Escucha los eventos `Alta`, `Editado` y `Borrado` de `Form_FormTecnico` (`:9`, `Form_FormTecnico.cls:8-10`). Declara `Eliminado` (`:5`) sin lanzarlo en este fichero.

## Controles y eventos

Controles: 32 (32 `Begin` en el `.form.txt`, sin contar `Begin Form`): 23 controles de usuario (filas siguientes), la sección `Detalle` (`.form.txt:286`) y 8 plantillas de estilo por defecto previas a la sección (`.form.txt:141-265`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaTecnicos` | lista de valores, 5 columnas con cabecera | `OnClick` → `ListaTecnicos_Click` | Muestra ALIAS, NOMBRE, TIPO, FECHAALTA, FECHABAJA. Al pulsar carga el `Tecnico` seleccionado y habilita los seis botones de acción | `.form.txt:496-559`; `:444-490` |
| `ComboTecnico` | combo de valores, 3 columnas | `AfterUpdate`; `Enter` (`Requery`) | Filtro por alias (contiene el texto) | `:374-402`; `.form.txt:1149-1172` |
| `ComboEnActivo` | combo de valores «Sí;No» | `AfterUpdate` | Filtro por baja: «Sí» = sin `FECHABAJA`, otro valor = con baja | `:349-372`; `.form.txt:734-755` |
| `ComandoLimpiarTecnico` / `ComandoLimpiarEnActivo` | botones (icono) | `OnClick` | Vacían su combo y refiltran | `:309-333`, `:284-308` |
| `ComandoAlta` | botón «Alta» | `OnClick` | Anula el técnico activo, cierra y abre `FormTecnico` | `:15-29`; `.form.txt:343-354` |
| `ComandoEdicion` | botón «Editar» (deshabilitado al inicio) | `OnClick` | Fija el técnico activo y abre `FormTecnico` | `:73-111`; `.form.txt:397-409` |
| `ComandoEliminar` | botón «Eliminar» (deshabilitado al inicio) | `OnClick` | Confirma y borra el técnico | `:157-206`; `.form.txt:425-438` |
| `ComandoAsignar` | botón «ASIGNAR» (deshabilitado al inicio) | **ninguno** | Se habilita con la selección pero no tiene `OnClick` ni handler | `.form.txt:640-666`; `:477-482` |
| `ComandoLibranzas` | botón «Libranzas» | `OnClick` | Abre `FormTecnicoAusenciasGestion` con el técnico activo | `:251-282`; `.form.txt:1186-1198` |
| `ComandoActividades` | botón «Actividades» | `OnClick` | Abre `FormActividadGestion` y asigna el alias a `ComboTécnicoBusqueda` (CAP-4) | `:34-70`; `.form.txt:1239-1251` |
| `ComandoInformeHoras` | botón «Informe Horas» | `OnClick` | Abre `FormTecnicoConsultaHoras` (CAP-10) | `:208-249`; `.form.txt:1266-1277` |
| `ComandoFiestas` | botón «Fiestas» | `OnClick` | Abre `FormTecnicoFiestasGestion` | `:113-128`; `.form.txt:1213-1224` |
| `ComandoTipoTecnico` | botón «Perfiles Técnicos» | `OnClick` | Abre `FormTipoTecnicoPreciosGestion` | `:335-347`; `.form.txt:1292-1303` |
| `ComandoInformeEnExcel` | botón (sin rótulo propio, `Caption` «Comando227») | `OnClick` | Exporta el contenido de la lista a Excel con `GenerarConsultas("Lista de Personal", …)` (CAP-10) | `:130-155`; `.form.txt:667-678` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra sin guardar | `:440-443`; `.form.txt:370-380` |
| `ComandoAyudaFormulario` | botón (icono, `Caption` «Comando46») | **ninguno** | Sin `OnClick` ni handler | `.form.txt:457-467` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta131` («Técnico»), `Etiqueta5` («En Activo») | etiquetas | — | Rótulos; `lblTitulo1` da título a la ventana | `.form.txt:613, 707, 768, 790`; `:417` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | `.form.txt:297, 319` |

## Flujos

### Apertura (`Form_Open`, `:404-439`)
1. `AjustarTamaño Me` (`:411`); título de ventana = `lblTitulo1` (`:417`).
2. `RellenarComboTecnicos Me.ComboTecnico, EnumSino.Sí` (`:418`; `FG.bas:7224-7260`) carga alias, nombre y tipo de **todos** los técnicos (`Constructor.bas:4264, 4275`).
3. `Filtrar` sin argumentos (`:423`): sin filtro de alias ni de baja. Si falla, `MsgBox` crítico y cierra el formulario (`:432-438`).

### Filtrado (`Filtrar`, `:492-578`)
Anula `m_ObjTecnicoActivo` y deshabilita los seis botones (`:510-517`), reinicia `RowSource` (`:519`), compone el `SELECT` y añade una fila por técnico (`:544-555`). Con una sola fila de datos (`ListCount = 2`, cabecera incluida) la selecciona y ejecuta `ListaTecnicos_Click` (`:558-560`); si no, solo lo ejecuta si hay selección (`:562-569`).

### Alta y edición
`ComandoAlta` anula el activo y abre `FormTecnico` (`:18-23`); `ComandoEdicion` lo fija al seleccionado (`:94-99`). Al guardar/borrar en `FormTecnico`, los handlers `m_Form_Alta`, `m_Form_Editado` y `m_Form_Borrado` recargan el combo y refiltran (`:580-645`).

### Eliminación (`:157-206`)
1. Sin selección: «Ha de seleccionar un elemento de la lista» (`:162-174`).
2. Confirmación `vbYesNo` con «No» por defecto (`:177-180`).
3. `Tecnico.Borrar` (`Tecnico.cls:392-413`) valida y ejecuta el `DELETE`; recarga combo y lista (`:182-196`).
4. Error: `MsgBox` crítico (`:200-205`).

### Exportación a Excel (`:130-155`)
`GenerarConsultas` recibe la lista (`:139`); recorre filas y columnas (`FG.bas:1443-1448`); exige consulta o lista, pero no ambas (`FG.bas:1413-1419`). Un `|` en el retorno se trata como error (`:140-144`).

## Reglas, validaciones y cálculos

- **Selección obligatoria** para editar, eliminar, libranzas, horas y actividades (`:42-55, 80-93, 162-175, 216-229, 256-269`).
- **Filtro de alias:** vacío ⇒ todos; si no, `ALIAS Like '*texto*'` (`:522-526`).
- **Filtro de baja:** vacío ⇒ todos; «Sí» ⇒ `FECHABAJA Is Null`; otro valor ⇒ `Not FECHABAJA Is Null` (`:527-534`).
- **Bloqueo de borrado:** no se elimina un técnico con registros en `TbEventos`, `TbActividades` o `TbGastos` (`Tecnico.cls:343-377`).
- **Control de botones:** seis botones solo habilitados con técnico seleccionado (`:449-454, 477-482`).
- **Error 1000:** convención interna para no sobrescribir mensajes ya fijados (`:26, 65`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado filtrado | `TbTecnicos` | SQL dinámico concatenado (alias entre comillas) vía `OpenRecordset` | — | `:520-544` |
| Carga del combo | `TbTecnicos` | Apertura de la tabla completa | — | `Constructor.bas:4275-4278` |
| Carga del técnico seleccionado | `TbTecnicos` | SQL dinámico `WHERE ALIAS='…'` | — | `Constructor.bas:1951-1953` |
| Comprobaciones previas al borrado | `TbEventos`, `TbActividades`, `TbGastos` | SQL dinámico con el alias concatenado (`Count`) | — | `Tecnico.cls:343-345, 357-359, 370-372` |
| Borrado | `TbTecnicos` | `DELETE` concatenado por `Execute` | no | `Tecnico.cls:403-405` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: sintaxis Jet/ACE `Like '*…*'` con comodín `*` (`:525, 523, 528`) y `DELETE *` (`Tecnico.cls:403`) que habrá que adaptar al dialecto de PostgreSQL; `Is Null` / `Not … Is Null` sobre `FECHABAJA`; alias como clave natural concatenada como texto en todos los accesos; sin transacciones.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoAsignar` se habilita y deshabilita con la selección (`:451, 479`) pero no tiene `OnClick` (`.form.txt:640-666`): botón sin función. Igual `ComandoAyudaFormulario` (`.form.txt:457-467`).
- `ComandoAlta_Click` no muestra el error al usuario: solo rellena `m_Error` (`:25-28`).
- `ComandoFiestas_Click` activa el reloj de arena (`:117`) y no lo desactiva en el camino correcto (`:120-121`).
- `Filtrar` concatena el texto del combo entre comillas (`:525`); un apóstrofo rompería la consulta. Igual `getTecnico` (`Constructor.bas:1951-1952`) y `Borrar` (`Tecnico.cls:403-404`).
- `Filtrar` recibe unas veces `ComboTecnico.Column(0)` (`:293, 357, 382`) y otras `ComboTecnico` (`:193, 589, 611`): podrían diferir si la columna enlazada no es la 0.
- `ListaTecnicos_Click` lanza `Err.Raise 100` (`:475`) y el manejador sustituye su mensaje por el genérico (`:485-486`).
- Si `m_ObjTecnicoSeleccionado` conserva un objeto antiguo, los botones no lo recargan (`:42, 80, 162`); depende de que la lista haya refrescado la selección.
- `ComandoActividades` asigna `ComboTécnicoBusqueda` por código (`:59`); a verificar si dispara su `AfterUpdate`.
- Rótulos de diseño inconsistentes: `Caption` «ComandoLimpiarEquipo» en los dos botones de limpieza (`.form.txt:823, 991`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro por parámetros enlazados en vez de SQL concatenado; clave sustituta en lugar del alias; borrado lógico con baja; definir o retirar «ASIGNAR».

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los técnicos y los botones de acción están deshabilitados (`:423`, `.form.txt:398, 426, 641, 1187, 1240, 1267`).
2. Seleccionar una fila habilita edición, eliminación, libranzas, horas y actividades.
3. Filtrar por alias muestra solo los que contienen el texto; «Sí»/«No» filtra por baja; los botones de limpieza restauran el filtro.
4. Si el filtro deja un único técnico, queda seleccionado.
5. Eliminar pide confirmación y no borra un técnico con eventos, actividades o gastos (mensaje de `Tecnico.cls:350, 364, 377`).
6. Tras crear, editar o borrar en `FormTecnico`, la lista y el combo se refrescan.
7. Exportar a Excel recoge las filas y columnas visibles en la lista.

## Pendientes

- Contenido de `AjustarTamaño` y de la plantilla/informe «Lista de Personal» usada por `GenerarConsultas` (`:139`).
- Columna enlazada de `ComboTecnico` (`.form.txt:1149-1172`) y de `ComboEnActivo`.
- Comprobación de perfil/permiso: no se ve en este `.cls`; a verificar en los formularios de origen (`Form_Form0BDOpciones.cls:112-115`).
- Longitud y unicidad de `ALIAS` en `TbTecnicos` (ERD de origen).
