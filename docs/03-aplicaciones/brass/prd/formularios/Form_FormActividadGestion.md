# Form_FormActividadGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadGestion.cls` (1065 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de actividades (`TbActividades`): lista filtrable por actividad, técnico y rango de fechas de alta, con acceso a alta, edición, eliminación, copia simple y múltiple (solo actividades que no son de evento), adjuntos, informe de horas del técnico y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista `ListaFiltrados` recibe en ejecución un `RowSource` SQL (`:1027`). Capacidad provisional: CAP-4 Actividades y tipos de actividad. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormActividadGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormActividadGestion"` | `Form_Form0BDOpciones.cls:90` |
| `Form_FormMaterialAltaReparacion` | si no está abierto lo abre, fija `ComboIDActividadBusqueda` y llama a `ComboIDActividadBusqueda_Change` | `Form_FormMaterialAltaReparacion.cls:267-276` |
| `Form_FormMaterialRepuestoAlta` | si no está abierto lo abre | `Form_FormMaterialRepuestoAlta.cls:82-83` |
| `Form_FormActividadAlta`, `Form_FormActividadDeEventoPorLotes` | `Requery` de `ListaFiltrados` si está abierto | `Form_FormActividadAlta.cls:80-81`; `Form_FormActividadDeEventoPorLotes.cls:287-288` |
| `Form_FormAnexos` | llama a `ListaFiltrados_Click` (público) para refrescar botones | `Form_FormAnexos.cls:75-76, 203-204` |

Abre: `FormActividadAlta` sin argumentos (`:249-254`); `FormActividadEdicion` con `OpenArgs` = `IDActividad|franqueado|facturado|URL anexo` (`:432-433`); `FormAnexos` tras fijar la global `m_ObjEntidadParaAnexoActiva` (`:224-226`); `FormTecnicoConsultaHoras` tras fijar `m_ObjTecnicoActivo` (`:567-571`). Cierra antes la instancia previa de cada uno con `FormularioAbierto` (`FG.bas:171-193`). `LimpiarUltimos` es pública y reinicia las memorias de filtro (`:175-180`).

## Controles y eventos

Controles: 29 con `Name` (38 Begin en el .form.txt: 1 `Form`, 7 bloques de valores por defecto, 1 `Section` y 29 controles).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista 8 columnas con encabezados | `OnClick` → `ListaFiltrados_Click` (público) | Columnas: 0 `IDActividad`, 1 descripción, 2 `HLab`, 3 `IDEvento`, 4 `Alta`, 5 `IDFactura`, 6 `Franqueado`, 7 `ALIASTECNICO`; habilita o no botones según selección | `:851-921`, `.form.txt:377-403` |
| `ComboIDActividadBusqueda` | combo 2 columnas | `OnChange`, `OnExit` | Filtra por actividad si cambia respecto a la memoria | `:635-681`, `.form.txt:1617-1644` |
| `ComboTécnicoBusqueda` | combo 3 columnas | `OnChange`, `OnExit` | Filtra por técnico; se rellena con `RellenarComboTecnicos` | `:683-730, 790`, `.form.txt:566-589` |
| `FechaInicialBusqueda`, `FechaFinalBusqueda` | texto (Short Date) | `OnExit` | Filtran por fecha de alta | `:732-777`, `.form.txt:683-726` |
| `ComandoLimpiarTecnico`, `ComandoLimpiarIDActividad`, `ComandoLimpiarFechas` | botones | `OnClick` | Vacían el filtro respectivo y relistan | `:588-633` |
| `ComandoAltaActividad` («Alta») | botón | `OnClick` | Abre `FormActividadAlta` | `:238-265` |
| `ComandoEditarActividad` («Editar») | botón (inicial deshabilitado) | `OnClick` | Abre `FormActividadEdicion` con la actividad | `:403-445` |
| `ComandoEliminarActividad` («Eliminar») | botón (inicial deshabilitado) | `OnClick` | Confirma y elimina actividad, materiales y anexos | `:446-504` |
| `ComandoCopiarActividad` («Copia Simple») | botón (inicial deshabilitado) | `OnClick` | Copia una actividad sin evento con fecha y horas por defecto | `:267-296` |
| `ComandoCopiarActividadMultiple` («Copia Múltiple») | botón (inicial deshabilitado) | `OnClick` | Copia en un rango de fechas laborables | `:298-401` |
| `ComandoAdjuntosActividad` («Adjuntar») | botón (inicial deshabilitado) | `OnClick` | Abre `FormAnexos` de la actividad | `:182-237` |
| `ImagenAnexoActividad` | imagen (oculta) | `OnClick` | Abre el anexo único en local, o `FormAnexos` si la URL contiene `;` | `:814-850` |
| `ComandoExportarExcelActividades` («Comando227») | botón (inicial deshabilitado) | `OnClick` | Exporta el resultado actual a Excel | `:505-548` |
| `ComandoInforme` («Informe horas») | botón | `OnClick` | Abre `FormTecnicoConsultaHoras` | `:550-585` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:809-812` |
| `lblFacturado` | etiqueta (oculta) | — | Visible si la actividad tiene `IDFactura` | `:883, 888`, `.form.txt:1712-1726` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta142`, `Etiqueta244`, `Etiqueta245`, `Etiqueta246` | etiquetas | — | Títulos y rótulos «TÉCNICO», «ENTRE», «Y», «ID.»; `lblTitulo1` da el título de ventana | `:789`, `.form.txt:903-1043` |
| `CuadroSuperior`, `Cuadro2`, `Marco122` | rectángulos y grupo de opciones | — | Decoración | `.form.txt:331-363, 1234-1244` |

## Flujos

### Apertura (`Form_Open`, `:778-808`)

`AjustarTamaño Me` (`FG.bas:261-287`), título de ventana, relleno del combo de técnicos con `RellenarComboTecnicos(…, EnumSino.Sí)` (`:787-790`; `FG.bas:7224`) y `ActualizarLista`. Si hay error, `MsgBox` y cierre del formulario (`:801-807`).

### Filtrado y lista (`Filtrar` `:922-994`, `ActualizarLista` `:995-1059`)

Sin ningún filtro la consulta devuelve `IDActividad='-1'` (lista vacía, `:952-957`). Con filtros, une `TbActividades` con `TbFacturaActividadesInvolucradas` y `TbEventos` (`LEFT JOIN`, `:945-951`) y ordena por `FechaAlta` descendente (`:984`). `ActualizarLista` deshabilita todos los botones, reasigna `RowSource`, hace `Requery` y, si hay una sola fila, la selecciona y llama a `ListaFiltrados_Click` (`:1006-1037`); `ComandoExportarExcelActividades` solo se habilita con filas (`:1031-1047`).

### Selección (`ListaFiltrados_Click`, `:851-921`)

Sin ID: deshabilita botones y oculta anexo y `lblFacturado` (`:861-875`). Con ID: habilita Editar y Adjuntar (`:876-877`); con `IDFactura` deshabilita Eliminar y muestra «Facturado» (`:884-889`); con `IDEvento` deshabilita las copias (`:890-896`). Relee la actividad con `Constructor.getActividad` y muestra el icono de anexo si `URLAnexos` no está vacía (`:897-907`).

### Eliminar (`:446-504`)

Sin selección: «Seleccione la actividad que pretende eliminar» (`:454-457`). Si `Franqueado` = «Sí»: error pidiendo quitar el franqueo (`:458-462`). Pide confirmación con «No» por defecto (`:471-476`) y llama a `ActividadEliminar(id, "Sí", "Sí", franqueado)` (`:485`), que borra anexos (`FG.bas:5686`), materiales (`FG.bas:5709`) y la actividad (`FG.bas:5739`). Después relista (`:494-496`).

### Copias (`:267-401`, `CopiarActividadNoDeEvento` `:8-89`)

Ambas exigen `IDEvento` vacío (`:276-279, 318-321`). La función copia todos los campos salvo `IDActividad`, `CodImportacion`, `IDFacturacion`, `IDExportacion`, `IDEVENTOGENERADO`, `FechaAlta` y `HorasLaborables` (`:58`), genera un ID con `DameIDActividad` (`:49`; `FG.bas:6535`), fija `FechaAlta` con `Format(…,"dd/mm/yyyy")` (`:64`) y `HorasLaborables` (por defecto «7,5», `:25-27, 65`). La copia múltiple pide dos fechas con `InputBox` (`:322-331`), valida orden (`:332-335`) y copia en cada día con horas registrables distintas de «0» (`:363-389`).

### Exportar a Excel (`:505-548`)

Recompone la consulta a partir del `RowSource` actual, parte por «Where» (`:515, 529-531`) y llama a `GenerarConsultasSQL` (`FG.bas:1584`), que genera un libro `.xlsx` en el directorio local del entorno mediante automatización de Excel (`FG.bas:1625`).

## Reglas, validaciones y cálculos

- **Horas registrables** (`HorasRegistrables`, `:90-169`): 0 en sábado y domingo (`:110-114`), en fecha de ausencia del técnico (`:115-126`), en fiesta (`:131-141`); si no, 7,5 menos las `HorasLaborables` de la actividad del técnico en esa fecha (`:147-162`).
- **Fechas del filtro:** sin inicial se usa 01/01/1900; sin final, 01/01/2100 (`:958-967`); se formatean `mm/dd/yyyy` entre `#` (`:961-979`).
- **Eliminación bloqueada** con factura (botón deshabilitado, `:887`; `ActividadEliminar` lo vuelve a rechazar, `FG.bas:5637-5640`) y con evento franqueado (`:458-462`).
- **Copia múltiple:** excluye días con 0 horas registrables (`:381`); el primer día se copia con la hora actual (`:364`).
- **Memorias de filtro** (`:7`): evitan relistar si el valor no cambia (`:643, 667, 691, 716, 740, 763`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de actividades | `TbActividades`, `TbFacturaActividadesInvolucradas`, `TbEventos` | SQL dinámico concatenado en `RowSource` | — | `:945-985` |
| Lectura de la actividad | `TbActividades` | SQL dinámico concatenado (`getActividad`) | — | `Constructor.bas:2294-2297` |
| Ausencias, fiestas, horas | `TbTecnicosAusencias`, `TbTecnicosFiestas`, `TbActividades` | SQL dinámico con fechas `#mm/dd/yyyy#` | — | `:115-150` |
| Copia de actividad | `TbActividades` | DAO `AddNew`/`Update` por campos | no | `:28-66` |
| ID nuevo | `TbActividades` | `DameValorSiguienteAlMaximoDeUnCampo` | no | `FG.bas:6535-6590` |
| Eliminación | `TbMaterial`, `TbActividades` y anexos | SQL dinámico + DAO `Delete` | no | `FG.bas:5693-5739` |
| Exportación | consulta de la lista | `OpenRecordset` + Excel | — | `:522-533`; `FG.bas:1584-1669` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literales de fecha `#mm/dd/yyyy#` (`:118, 133, 150, 961-979`), `Like '*'` (`:970, 975`), joins entre paréntesis (`:949-951`), campo `DESCRIPCIÓN` con tilde como alias (`:945`), `HorasLaborables` asignada desde cadena con coma decimal (`:26, 65`), IDs de actividad textuales calculados en VBA con máximo+1, borrado en cascada manual sin transacción y exportación por Excel local.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `CopiarActividadNoDeEvento` en error asigna `p_Error` pero no su valor de retorno (`:76-88`); en la copia múltiple se llama sin `p_Error` y se busca `|` en el retorno (`:343-348, 365-370, 382-387`), por lo que un fallo de copia pasaría sin aviso.
- `HorasRegistrables` lee solo la primera actividad del técnico en esa fecha (`:153-157`); con varias actividades el resto se ignora. La jornada de 7,5 está fija (`:161`) y puede dar valor negativo, que se trata como registrable al no ser «0» (`:363, 381`).
- `ComandoEliminarActividad_Click` pasa el valor de franqueo como `strEstaFacturado` (`:485`; `FG.bas:5581-5586`); un «No» evitaría la comprobación de factura en `ActividadEliminar`.
- El resultado de `ActividadEliminar` (materiales no eliminados) no se muestra (`:485-496`); su doble comprobación de `|` hace inalcanzable la rama de no eliminados (`FG.bas:5714-5715`).
- Eliminación y copia sin transacción: un fallo intermedio deja materiales o anexos borrados con la actividad viva (`FG.bas:5686-5739`).
- Exportación: `Split` por «Where» tomando `dato(1)` (`:529-530`) perdería texto si «Where» apareciese dos veces; el filtro por ID y alias concatena sin escapar (`:972, 977`).
- `LEFT JOIN` con facturas (`:949-951`) podría duplicar filas si una actividad figura en varias facturas.
- `ComandoInforme_Click` abre el informe aun sin técnico (`m_ObjTecnicoSeleccionado` siempre `Nothing` al inicio, `:551-571`) y llama `Hourglass False` en lugar de `True` (`:554`).
- `Franqueado` se compara con «Sí» (`:459`) sin confirmar cómo muestra la lista el valor lógico.
- `ListaFiltrados` declara 8 columnas y 9 anchos (`.form.txt:383, 402`); `ComboIDActividadBusqueda` conserva un `RowSource` de diseño truncado (`.form.txt:1640`).
- Variables sin uso en `ComandoAdjuntosActividad_Click` (`strEscritura`, `:204-208`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: transacción única para eliminar y copiar; jornada y festivos parametrizados; resultado de borrado visible; consultas parametrizadas; exportación en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin ningún filtro la lista queda vacía (`:952-957`).
2. Filtrar por técnico, actividad o fechas devuelve solo coincidencias, de más reciente a más antigua (`:969-984`).
3. Con una actividad facturada, Eliminar está deshabilitado y se ve «Facturado» (`:884-889`).
4. Eliminar una actividad de evento franqueado muestra el aviso de quitar franqueo (`:458-462`).
5. Confirmar la eliminación borra la actividad, sus materiales y anexos y relista (`:471-496`).
6. Las copias están deshabilitadas para actividades de evento (`:890-893`) y rechazan su uso (`:276-279`).
7. La copia múltiple no crea actividades en fin de semana, ausencia ni fiesta (`:110-141, 381`).
8. Un periodo final anterior al inicial muestra «Periodo de fechas incorrecto» (`:332-335`).
9. Exportar con filas genera el libro Excel; sin filas, el botón está deshabilitado (`:1031-1047`).

## Pendientes

- Contenido de `Constructor.getTecnico` y de la clase `Actividad` (`:209, 559, 901`).
- Efecto real de `EliminarAnexos` y `EliminarMaterialEnActividad` (`FG.bas:5686, 5709`).
- Estructura de `TbActividades` y claves de `TbFacturaActividadesInvolucradas`: ERD de origen.
- Qué hace `Ejecutar` (`FG.bas:246`) y cómo resuelve el directorio local el entorno.
