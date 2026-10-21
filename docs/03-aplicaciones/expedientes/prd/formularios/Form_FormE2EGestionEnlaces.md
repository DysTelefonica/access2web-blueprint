# Form_FormE2EGestionEnlaces

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormE2EGestionEnlaces.cls` (261 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Consola de administración de los enlaces E2E: lista las familias de expedientes enlazadas (`TbE2EEnlaces`), permite desvincular una familia completa, encolar su exportación a E2E, encolar todas las pendientes y generar el JSON de una familia en una carpeta local. Es un formulario **no enlazado**: la lista es un `Value List` que se rellena por código (`.form.txt:640-660`; `FormE2EGestionEnlacesLogic.bas:23-70`). El `.cls` solo delega en `FormE2EGestionEnlacesLogic.bas` (`:5-9`). Las citas son `Form_FormE2EGestionEnlaces.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` | Opción 11 («Gestor E2E») de `ListaOtrasOpciones`, resuelta a `FormE2EGestionEnlaces` y abierta con `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:33-34`, `Form_Form0BDGestorEntidades.cls:58-71`, `classes/Entorno.cls:700, 739` |
| `Form_Form0BDOpciones` | `ComandoGestionE2EEnlaces_Click` cierra el formulario si estaba abierto y lo abre | `Form_Form0BDOpciones.cls:151-161` |

No abre otros formularios. Lanza un selector de carpeta (`Application.FileDialog`, `:200-205`) y escribe un fichero JSON (`:211-215`). Sale con `cmdSalir` sin guardar (`:13-16`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ModoHistorico` | combo (`Activos;Histórico;Todos`) | `AfterUpdate` | Relanza la búsqueda | `:18-27`, `.form.txt:586-589` |
| `PalabraClave` | texto | `AfterUpdate` | Relanza la búsqueda | `:40-49`, `.form.txt:873-874` |
| `ComandoBuscar` | botón | `OnClick` | Relanza la búsqueda | `:29-38` |
| `ListaEnlaces` | lista, 7 columnas, ID oculto, con cabecera | `OnClick` | Habilita los comandos según la fila y muestra «Familia: N hijos» | `:55-74`, `.form.txt:641-660` |
| `ComandoDesvincular` | botón (deshabilitado en diseño) | `OnClick` | Da de baja el enlace de toda la familia | `:76-111`, `.form.txt:679, 689-691` |
| `ComandoExportarE2E` («Exportar a E2E ahora») | botón (deshabilitado en diseño) | `OnClick` | Encola una instantánea de la familia en `TbE2ECola` | `:113-144`, `.form.txt:717, 727-729` |
| `ComandoExportarTodosPendientes` («Exportar todos») | botón (deshabilitado en diseño) | `OnClick` | Encola todos los enlazados con cambios pendientes | `:152-183`, `.form.txt:933, 943-946` |
| `ComandoGenerarJSON` | botón (deshabilitado en diseño) | `OnClick` | Genera el JSON de la familia y lo guarda en una carpeta elegida | `:185-228`, `.form.txt:895, 905-907` |
| `lblFamiliaCount` | etiqueta (oculta en diseño) | — | «Familia: N hijos» tras seleccionar | `.form.txt:1057, 1065-1066`; `FormE2EGestionEnlacesLogic.bas:166-167` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:13-16` |

## Flujos

### Apertura (`Form_Open`, `:230-261`)

1. `Ajustar`, título desde la global `m_TituloFormulario`, tres comandos deshabilitados y lista reducida a la fila de cabecera `LISTA_ENLACES_HEADERS` (`:234-239`; `FormE2EGestionEnlacesLogic.bas:13`).
2. `ValidarAcceso` con `EsAdministrador = Sí`; si no lo es, `Cancel = True` y aviso «Solo Administradores o Gestores pueden acceder…» (`:241-246`; `FormE2EGestionEnlacesLogic.bas:503-511`). `EsAdministrador` sale de `EsAdministradorCalculado` (`Variables Globales.bas:383`).
3. Filtros por defecto `Activos` y palabra vacía, y primera búsqueda (`:248-252`).
4. Cualquier error muestra `MsgBox` crítico y cierra el formulario (`:254-260`).

### Búsqueda (`EjecutarBusqueda`, `FormE2EGestionEnlacesLogic.bas:342-359`)

1. Construye el SQL sobre `TbE2EEnlaces INNER JOIN TbExpedientes` (`:455-459`) con el filtro de modo: Activos ⇒ `FechaBaja IS NULL`; Histórico ⇒ `IS NOT NULL`; Todos ⇒ sin filtro (`:461-468`).
2. Solo raíces de familia: `IDExpedientePadre IS NULL` en todos los modos (`:470-474`).
3. Con palabra clave: `InStr` en `Titulo` o `Nemotecnico`, igualdad en `CodExp` o igualdad numérica en `IDExpediente` (`:476-484`). Orden `FechaAlta DESC` (`:486-487`).
4. Vacía la lista reasignando la cabecera, añade una fila por registro con `AddItem` y deshabilita los tres comandos (`:43-63`). Fechas en `dd/mm/yyyy` o vacías (`:322-331`).

### Selección (`ListaEnlaces_Click`, `:55-74`)

Lee `Column(0)` (ID) y `Column(5)` (fecha de baja) y llama a `AplicarSeleccion` (`:61-63`): Desvincular y Exportar solo con ID > 0 y sin fecha de baja; Generar JSON con cualquier ID > 0 (`FormE2EGestionEnlacesLogic.bas:139-144, 586-601`). Cuenta los hijos directos con `COUNT(*)` sobre `IDExpedientePadre` (`:644-662`). Si falla, deshabilita los tres comandos (`:71-73`).

### Desvincular (`:76-111`)

1. Exige ID seleccionado: «Debe seleccionar un expediente activo de la lista» (`:83-87`).
2. `DesvincularFamilia` con motivo fijo «Baja desde form de Gestión» y el usuario conectado (`:96-98`; `FormE2EGestionEnlacesLogic.bas:627-628`).
3. Dentro de una transacción de `Workspaces(0)`: descubre la familia, y para cada miembro vinculado pone `FechaBaja = Now()`, `MotivoBaja`, `IDUsuarioBaja` y `OrdinalE2E = Null` en `TbExpedientes`; después recalcula la caché de entidades con ámbito `E2E` de cada miembro (`E2EEnlacesService.bas:191-218`). Error ⇒ `Rollback` (`:222-230`).
4. Mensaje de éxito y nueva búsqueda (`:101-104`).

### Exportar a E2E ahora (`:113-144`)

`EncolarExportacionE2E` (`FormE2EGestionEnlacesLogic.bas:180-210`) → `E2EQueue_EnqueueFamilySnapshot` con origen «Manual» (`ManualE2EExport.bas:55-56`): exige enlace activo («Expediente no vinculado a E2E», `:98-101`), toma la raíz de la familia, calcula `OrdinalE2E` = máximo + 1, genera el JSON canónico y su hash e inserta una fila `Pendiente` en `TbE2ECola` (`:103-138`). Éxito: mensaje «JSON generado y encolado…» (`FormE2EGestionEnlacesLogic.bas:202`) y nueva búsqueda (`:135-137`).

### Exportar todos (`:152-183`)

`ExportarTodosPendientesE2E` (`FormE2EGestionEnlacesLogic.bas:227-296`): lista los expedientes enlazados activos con `HashActual` distinto del último `Enviado` de la cola (`ManualE2EExport.bas:238-279`); sin pendientes avisa «No hay expedientes pendientes de exportar.» (`:316-318`); con pendientes pide confirmación con hasta 10 `CodExp` (`:311-337`), encola uno a uno sin abortar ante fallos (`:437-451`) y muestra «Encolados / Fallidos» (`:345-351`). Refresca la lista si encoló alguno (`:173-176`).

### Generar JSON (`:185-228`)

Exige ID (`:194-198`), pide carpeta («Exportación cancelada» si se cancela, `:200-205`), genera el JSON con `GenerarJsonE2ECanonicoPorListaConFamiliaParaBatch(0, ID)` (`FormE2EGestionEnlacesLogic.bas:390-392`) y lo escribe como `E2E_Expediente_<ID>_<aaaammddhhnnss>.json` con `ADODB.Stream` UTF-8 (`:211-215`; `FormE2EGestionEnlacesLogic.bas:421-428`). Es el único handler que envía `CorreoAlAdministrador` (`:220-226`).

## Reglas, validaciones y cálculos

- Acceso: solo `EsAdministrador = Sí` (`:241`; `FormE2EGestionEnlacesLogic.bas:510`).
- La lista muestra solo raíces; desvincular o exportar actúa siempre sobre la familia completa (`FormE2EGestionEnlacesLogic.bas:474, 623-628`; `ManualE2EExport.bas:103-113`).
- Un enlace dado de baja no puede desvincularse ni encolarse; sí puede generar JSON (`FormE2EGestionEnlacesLogic.bas:97-113, 586-601`).
- Escape de la palabra clave: comillas duplicadas y `*`, `?` entre corchetes (`FormE2EGestionEnlacesLogic.bas:476-478`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de enlaces | `TbE2EEnlaces`, `TbExpedientes` | SQL dinámico, snapshot | No | `FormE2EGestionEnlacesLogic.bas:455-487, 46` |
| Contar hijos | `TbExpedientes` | SQL dinámico | No | `FormE2EGestionEnlacesLogic.bas:657-659` |
| Desvincular | `UPDATE TbE2EEnlaces`, `UPDATE TbExpedientes.OrdinalE2E` | SQL dinámico `Execute` | Sí, `Workspaces(0)` | `E2EEnlacesService.bas:191-218` |
| Encolar | `INSERT INTO TbE2ECola` | SQL dinámico `Execute` | No | `ManualE2EExport.bas:132-138` |
| Pendientes | `TbExpedientes`, `TbE2EEnlaces`, `TbE2ECola` | SQL dinámico, N+1 | No | `ManualE2EExport.bas:238-255` |
| JSON local | fichero en carpeta elegida | `ADODB.Stream` | — | `FormE2EGestionEnlacesLogic.bas:421-428` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `InStr`, `Now()`, `TOP 1` y `1=1` en SQL Jet (`FormE2EGestionEnlacesLogic.bas:465, 480-481`; `ManualE2EExport.bas:250`); transacción DAO de espacio de trabajo con `Execute` (`E2EEnlacesService.bas:191-218`); el ordinal de cola se calcula como máximo + 1 en VBA y una colisión da el error 3022 (`ManualE2EExport.bas:115, 144-145`); el JSON entero viaja como literal en el `INSERT` (`:137`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoExportarTodosPendientes` está deshabilitado en el diseño (`.form.txt:933`) y ningún código lo habilita: el botón no sería pulsable. Además, el código embebido en el `.form.txt` no contiene su handler (`.form.txt:1076-1300` frente a `:146-184`).
- La columna «Motivo» de la lista siempre sale vacía: el SQL lee `MotivoBaja` pero `BuildValueListRow` solo compone 6 columnas (`FormE2EGestionEnlacesLogic.bas:13, 457, 333-338`).
- Un `;` en `Titulo` o `Nemotecnico` desplazaría las columnas del `Value List`, y `Column(5)` dejaría de ser la fecha de baja (`FormE2EGestionEnlacesLogic.bas:333-338`; `:62`).
- Si `getdb` falla o el SQL da error, la lista queda vacía o sin cambios sin mostrar error (`FormE2EGestionEnlacesLogic.bas:35-36, 66-69`).
- El escape `[*]`/`[?]` es de `LIKE`, pero la búsqueda usa `InStr` e igualdad: buscar un texto con `*` o `?` no encontraría nada (`FormE2EGestionEnlacesLogic.bas:477-483`).
- En `ModoHistorico`, `ComandoBuscar` y `PalabraClave`, el mensaje del módulo se sobrescribe con el de `Err.Raise 1000` y se pierde (`:22-25, 33-36, 44-47`).
- «Familia: N hijos» cuenta solo hijos directos, mientras que desvincular y exportar recorren toda la descendencia (`FormE2EGestionEnlacesLogic.bas:657-659`; `E2EEnlacesService.bas:194`).
- El mensaje de desvincular ignora las filas afectadas (`:96-102`); además el contador suma `RecordsAffected` del `UPDATE` de `TbExpedientes`, no del de `TbE2EEnlaces` (`E2EEnlacesService.bas:199-205`).
- «Exportar todos» lista también descendientes enlazados y cada uno encola la instantánea de su raíz: una familia podría encolarse varias veces (`ManualE2EExport.bas:238-246, 113-117`).
- Los avisos informativos de «Exportar todos» (sin pendientes, resumen) se muestran con icono crítico (`FormE2EGestionEnlacesLogic.bas:266-267, 290`; `modUiSuppress.bas:22`).
- El comentario de `GuardarJSONEn` promete quitar el BOM, pero el código no lo hace (`FormE2EGestionEnlacesLogic.bas:401-406, 416-428`).
- El aviso de acceso nombra a «Gestores», pero solo se comprueba `EsAdministrador` (`:243`; `FormE2EGestionEnlacesLogic.bas:510`). Al cancelar `Form_Open`, `DoCmd.OpenForm` del llamador podría lanzar el error 2501 y enviar correo al administrador (`Form_Form0BDGestorEntidades.cls:71, 82-88`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: lista paginada con columnas tipadas en lugar de `Value List`; confirmación antes de desvincular una familia; descarga del JSON por navegador en lugar de carpeta local; ordinal de cola por secuencia de base de datos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador no accede a la pantalla.
2. Al abrir, la lista muestra solo raíces enlazadas activas, ordenadas por fecha de alta descendente.
3. Seleccionar un enlace dado de baja deja deshabilitados Desvincular y Exportar, y habilitado Generar JSON.
4. Desvincular da de baja todos los miembros activos de la familia y vacía su `OrdinalE2E` en una sola transacción.
5. Exportar un enlace activo crea una fila `Pendiente` en la cola asociada a la raíz.
6. Exportar todos sin pendientes informa «No hay expedientes pendientes de exportar.» y no escribe.

## Pendientes

- Si el control `ComandoGestionE2EEnlaces` existe en `Form_Form0BDOpciones`: en su `.form.txt` solo aparece en el código embebido (`Form_Form0BDOpciones.form.txt:3305, 3314`).
- Proceso que copia la cola al directorio vigilado «en el siguiente ciclo de 5 min» (`FormE2EGestionEnlacesLogic.bas:202`).
- Contenido de `GenerarJsonE2ECanonicoPorListaConFamiliaParaBatch` (`ExpedienteJsonExporter.bas:358`), `DescubrirFamilia` (`E2EEnlacesService.bas:30`) y `ExpedienteEntidadOperaciones.Registrar` con ámbito `E2E`.
- Si la base de `getdb` pertenece a `Workspaces(0)`, condición para que la transacción de desvincular la cubra (`Variables Globales.bas:607`; `E2EEnlacesService.bas:191`).
- Quién asigna `m_TituloFormulario` (`:235`) y qué hace `Ajustar` (`FUNCIONES UTILES.bas:2703`).
