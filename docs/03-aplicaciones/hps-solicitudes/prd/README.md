# PRD — HPS Solicitudes

> **Estado:** fase 1 (fijación e inventario). Los formularios, módulos y clases quedan **pendientes**; las fichas, la matriz, la coexistencia, las capacidades y las issues propuestas llegan en las PR siguientes de la cadena.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/HPS_SOLICITUDES |
| Commit fijado | `cfa1937ab05a9e40150ea9236243898de96f84da` (`main`, 2026-04-16, «chore: hard reset main - code VBA con fix + .accdb + .ini (sin deploy.ps1)») |
| Índice CodeGraph-VBA | 108 ficheros indexados, 5 653 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md) (DRAFT), walkthroughs `walkthrough-*.json`, [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/hps-solicitudes/src` (2026-08-06) |
| Issue | #1077 (método y plantilla de #1069) |

## Alcance actual

Fija el commit y concilia el inventario. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `cfa1937` | Snapshot | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/Form_*.cls`, con su `.form.txt`) | 34 | 34 | +0 |
| Módulos estándar (`src/modules/*.bas`) | 14 | 14 | +0 |
| Clases (`src/classes/*.cls`) | 26 | 26 | +0 |

La comparación normaliza las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas. Diferencia textual cero no demuestra que el estudio previo sea correcto: solo que el código no ha cambiado desde el snapshot.

### Delta de código frente al snapshot

- **Distintos del snapshot:** 0 formularios, 0 módulos y 0 clases.
- **Nuevos en el commit (sin equivalente en el snapshot):** 0 formularios, 0 módulos y 0 clases.
- **Del snapshot ausentes en el commit:** 0 formularios, 0 módulos y 0 clases.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub`, `Function` y `Property` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide tras la normalización.

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Form_Form0BDOpciones | src/forms/Form_Form0BDOpciones.cls | 407 | 33 | 13 | igual | analizado | [formularios/Form_Form0BDOpciones.md](formularios/Form_Form0BDOpciones.md) |
| Form_FormAdjuntaTraspasoONS | src/forms/Form_FormAdjuntaTraspasoONS.cls | 234 | 28 | 7 | igual | pendiente | — |
| Form_FormAdjuntarExcelSolicitante | src/forms/Form_FormAdjuntarExcelSolicitante.cls | 215 | 28 | 7 | igual | analizado | [formularios/Form_FormAdjuntarExcelSolicitante.md](formularios/Form_FormAdjuntarExcelSolicitante.md) |
| Form_FormAltaDatosSolicitante | src/forms/Form_FormAltaDatosSolicitante.cls | 452 | 63 | 15 | igual | pendiente | — |
| Form_FormCambioTipo | src/forms/Form_FormCambioTipo.cls | 107 | 26 | 4 | igual | pendiente | — |
| Form_FormConfiguracion | src/forms/Form_FormConfiguracion.cls | 418 | 50 | 11 | igual | analizado | [formularios/Form_FormConfiguracion.md](formularios/Form_FormConfiguracion.md) |
| Form_FormCorreo | src/forms/Form_FormCorreo.cls | 403 | 36 | 10 | igual | pendiente | — |
| Form_FormCorreoPruebas | src/forms/Form_FormCorreoPruebas.cls | 259 | 38 | 8 | igual | pendiente | — |
| Form_FormCorreosGestion | src/forms/Form_FormCorreosGestion.cls | 541 | 34 | 19 | igual | pendiente | — |
| Form_FormEntidades | src/forms/Form_FormEntidades.cls | 170 | 17 | 7 | igual | analizado | [formularios/Form_FormEntidades.md](formularios/Form_FormEntidades.md) |
| Form_FormEnvioCorreo | src/forms/Form_FormEnvioCorreo.cls | 428 | 36 | 10 | igual | pendiente | — |
| Form_FormExpedienteDetalle | src/forms/Form_FormExpedienteDetalle.cls | 79 | 32 | 3 | igual | analizado | [formularios/Form_FormExpedienteDetalle.md](formularios/Form_FormExpedienteDetalle.md) |
| Form_FormExpedientesBusqueda | src/forms/Form_FormExpedientesBusqueda.cls | 352 | 27 | 10 | igual | analizado | [formularios/Form_FormExpedientesBusqueda.md](formularios/Form_FormExpedientesBusqueda.md) |
| Form_FormMotivoHPS | src/forms/Form_FormMotivoHPS.cls | 124 | 23 | 4 | igual | pendiente | — |
| Form_FormMotivosHPS | src/forms/Form_FormMotivosHPS.cls | 323 | 18 | 10 | igual | pendiente | — |
| Form_FormPlantillasHTML | src/forms/Form_FormPlantillasHTML.cls | 181 | 17 | 7 | igual | pendiente | — |
| Form_FormResponsable | src/forms/Form_FormResponsable.cls | 214 | 28 | 6 | igual | pendiente | — |
| Form_FormResponsablesGestion | src/forms/Form_FormResponsablesGestion.cls | 392 | 28 | 13 | igual | pendiente | — |
| Form_FormSolicitudAlta | src/forms/Form_FormSolicitudAlta.cls | 130 | 24 | 4 | igual | analizado | [formularios/Form_FormSolicitudAlta.md](formularios/Form_FormSolicitudAlta.md) |
| Form_FormSolicitudAltaTraspaso | src/forms/Form_FormSolicitudAltaTraspaso.cls | 142 | 25 | 5 | igual | pendiente | — |
| Form_FormSolicitudRenovacion | src/forms/Form_FormSolicitudRenovacion.cls | 151 | 25 | 5 | igual | pendiente | — |
| Form_FormSolicitudRenovacionDatos | src/forms/Form_FormSolicitudRenovacionDatos.cls | 957 | 64 | 19 | igual | pendiente | — |
| Form_FormSolicitudRenovacionFechas | src/forms/Form_FormSolicitudRenovacionFechas.cls | 1652 | 82 | 40 | igual | pendiente | — |
| Form_FormSolicitudesAltaFechas | src/forms/Form_FormSolicitudesAltaFechas.cls | 1661 | 82 | 40 | igual | analizado | [formularios/Form_FormSolicitudesAltaFechas.md](formularios/Form_FormSolicitudesAltaFechas.md) |
| Form_FormSolicitudesAltaRenovacionFechas | src/forms/Form_FormSolicitudesAltaRenovacionFechas.cls | 751 | 63 | 18 | igual | pendiente | — |
| Form_FormSolicitudesAltaTraspasoDatos | src/forms/Form_FormSolicitudesAltaTraspasoDatos.cls | 953 | 66 | 23 | igual | pendiente | — |
| Form_FormSolicitudesGestion | src/forms/Form_FormSolicitudesGestion.cls | 912 | 48 | 23 | igual | analizado | [formularios/Form_FormSolicitudesGestion.md](formularios/Form_FormSolicitudesGestion.md) |
| Form_FormSolicitudesTraspasoFechas | src/forms/Form_FormSolicitudesTraspasoFechas.cls | 1166 | 56 | 29 | igual | pendiente | — |
| Form_FormTareasTramitadorPendientes | src/forms/Form_FormTareasTramitadorPendientes.cls | 652 | 29 | 14 | igual | pendiente | — |
| Form_FormUsuariosHPSBusqueda | src/forms/Form_FormUsuariosHPSBusqueda.cls | 295 | 26 | 9 | igual | pendiente | — |
| Form_FormWeb | src/forms/Form_FormWeb.cls | 56 | 16 | 2 | igual | analizado | [formularios/Form_FormWeb.md](formularios/Form_FormWeb.md) |
| Form_frmBuscarJustificacion | src/forms/Form_frmBuscarJustificacion.cls | 333 | 29 | 17 | igual | analizado | [formularios/Form_frmBuscarJustificacion.md](formularios/Form_frmBuscarJustificacion.md) |
| Form_frmJustificacion | src/forms/Form_frmJustificacion.cls | 139 | 28 | 6 | igual | analizado | [formularios/Form_frmJustificacion.md](formularios/Form_frmJustificacion.md) |
| Form_frmSplash | src/forms/Form_frmSplash.cls | 63 | 15 | 2 | igual | analizado | [formularios/Form_frmSplash.md](formularios/Form_frmSplash.md) |

### Módulos estándar

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| AutomatizacionRepositorio | src/modules/AutomatizacionRepositorio.bas | 61 | — | 1 | igual | pendiente | — |
| Automiatizacion | src/modules/Automiatizacion.bas | 215 | — | 3 | igual | pendiente | — |
| CorreoRepositorio | src/modules/CorreoRepositorio.bas | 115 | — | 2 | igual | pendiente | — |
| FUNCIONES UTILES | src/modules/FUNCIONES UTILES.bas | 2544 | — | 64 | igual | pendiente | — |
| Factoria | src/modules/Factoria.bas | 19 | — | 1 | igual | pendiente | — |
| InformesOperaciones | src/modules/InformesOperaciones.bas | 336 | — | 6 | igual | pendiente | — |
| Instalador | src/modules/Instalador.bas | 146 | — | 3 | igual | pendiente | — |
| JsonConverter | src/modules/JsonConverter.bas | 1401 | — | 31 | igual | pendiente | — |
| JustificacionRepositorio | src/modules/JustificacionRepositorio.bas | 107 | — | 6 | igual | pendiente | — |
| RepositorioComun | src/modules/RepositorioComun.bas | 95 | — | 6 | igual | pendiente | — |
| Test | src/modules/Test.bas | 386 | — | 11 | igual | pendiente | — |
| TestParametrosParser | src/modules/TestParametrosParser.bas | 281 | — | 11 | igual | pendiente | — |
| Variables Globales | src/modules/Variables Globales.bas | 698 | — | 11 | igual | pendiente | — |
| constructor | src/modules/constructor.bas | 3922 | — | 63 | igual | pendiente | — |

### Clases

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Configuracion | src/classes/Configuracion.cls | 241 | — | 4 | igual | pendiente | — |
| ConfiguracionOperaciones | src/classes/ConfiguracionOperaciones.cls | 402 | — | 3 | igual | pendiente | — |
| Correo | src/classes/Correo.cls | 374 | — | 7 | igual | pendiente | — |
| CorreoOperaciones | src/classes/CorreoOperaciones.cls | 287 | — | 4 | igual | pendiente | — |
| CorreoServicio | src/classes/CorreoServicio.cls | 34 | — | 1 | igual | pendiente | — |
| Entorno | src/classes/Entorno.cls | 3065 | — | 115 | igual | pendiente | — |
| Expediente | src/classes/Expediente.cls | 554 | — | 5 | igual | pendiente | — |
| Justificacion | src/classes/Justificacion.cls | 57 | — | 10 | igual | pendiente | — |
| JustificacionServicio | src/classes/JustificacionServicio.cls | 114 | — | 7 | igual | pendiente | — |
| LogGeneral | src/classes/LogGeneral.cls | 265 | — | 7 | igual | pendiente | — |
| LogGeneralOperaciones | src/classes/LogGeneralOperaciones.cls | 69 | — | 1 | igual | pendiente | — |
| MotivoHPS | src/classes/MotivoHPS.cls | 333 | — | 10 | igual | pendiente | — |
| ParametrosParser | src/classes/ParametrosParser.cls | 213 | — | 5 | igual | pendiente | — |
| Responsable | src/classes/Responsable.cls | 137 | — | 4 | igual | pendiente | — |
| ResponsableOperaciones | src/classes/ResponsableOperaciones.cls | 166 | — | 3 | igual | pendiente | — |
| Solicitud | src/classes/Solicitud.cls | 1744 | — | 50 | igual | pendiente | — |
| SolicitudFechas | src/classes/SolicitudFechas.cls | 789 | — | 19 | igual | pendiente | — |
| SolicitudFechasOperaciones | src/classes/SolicitudFechasOperaciones.cls | 532 | — | 6 | igual | pendiente | — |
| SolicitudOperaciones | src/classes/SolicitudOperaciones.cls | 2653 | — | 35 | igual | pendiente | — |
| Suministrador | src/classes/Suministrador.cls | 187 | — | 4 | igual | pendiente | — |
| UltimoCambio | src/classes/UltimoCambio.cls | 232 | — | 7 | igual | pendiente | — |
| UltimoCambioOperaciones | src/classes/UltimoCambioOperaciones.cls | 93 | — | 1 | igual | pendiente | — |
| Usuario | src/classes/Usuario.cls | 570 | — | 11 | igual | pendiente | — |
| UsuarioAplicacionPermisos | src/classes/UsuarioAplicacionPermisos.cls | 368 | — | 11 | igual | pendiente | — |
| UsuarioEntidad | src/classes/UsuarioEntidad.cls | 1099 | — | 3 | igual | pendiente | — |
| UsuarioHPS | src/classes/UsuarioHPS.cls | 409 | — | 7 | igual | pendiente | — |

Totales: 13 analizados · 0 no aplicables · 61 pendientes.



### Hallazgos del legado (fase 1)

Comportamientos sospechosos detectados al analizar los formularios. No se corrigen, no generan issues, y cada uno queda **a verificar** en ejecución.

| Id | Hallazgo | Evidencia | Ficha |
|---|---|---|---|
| H-HPS-S-1 | Splash hereda `Caption = "CONDOR - Iniciando..."` del diseño y solo se sobreescribe al disparar `OnTimer`; la ventana arranca con un rótulo de otra aplicación. | `Form_frmSplash.form.txt:40`; `Form_frmSplash.cls:50` | [Form_frmSplash](formularios/Form_frmSplash.md) |
| H-HPS-S-2 | Subtítulo «Control y Gestión de las Solcitudes de HPS» con typo «Solcitudes» en el subtítulo fijo. | `Form_frmSplash.form.txt:1259` | [Form_frmSplash](formularios/Form_frmSplash.md) |
| H-HPS-S-3 | `/dev` se pasa también a `EVE` como pseudo-correo, así que el arranque cae en «usuario no registrado». | `Form_frmSplash.cls:17`; `VG.bas:103-108` | [Form_frmSplash](formularios/Form_frmSplash.md) |
| H-HPS-S-4 | La rama `errores:` de `Form_Timer` abre `Form0BDOpciones` para no cerrar la app aunque `EVE` falle: deja entrar al menú con la identidad local de Windows. | `Form_frmSplash.cls:74-79` | [Form_frmSplash](formularios/Form_frmSplash.md) |
| H-HPS-S-5 | El botón «Ayuda» del menú principal envía un correo real al administrador con la cadena «Prueba de Error» antes de llamar a `AbrirAyuda`; marcador residual. | `Form_Form0BDOpciones.cls:83-95` | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) |
| H-HPS-S-6 | Para perfil técnico, el contador de `lblTareasPendientes` siempre muestra `0` literal en lugar del total por usuario (`0/<global>`). | `Form_Form0BDOpciones.cls:393-405` | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) |
| H-HPS-S-7 | `m_NombreTramitadorAlInicio` se declara en `Form_Form0BDOpciones.cls:5` y nunca se usa; candidata a retirar o referencia dinámica no resuelta. | `Form_Form0BDOpciones.cls:5` | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) |
| H-HPS-S-8 | `ComandoTareasPendientes_Click` re-engancha `m_FormTareas` cada vez sin `Set m_FormTareas = Nothing` previo; los `WithEvents` podrían mantener referencias colgantes. | `Form_Form0BDOpciones.cls:265-269` | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md) |
| H-HPS-S-9 | Los cuatro `errores:` de `Form_FormConfiguracion` repiten el literal «Al Form_Load se ha producido» aunque estén en `ComandoAyuda_Click`, `ComandoRegistrar_Click` o `ComandoURLPlantillas*_Click`. | `Form_FormConfiguracion.cls:42, 84, 218, 280` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) |
| H-HPS-S-10 | `cmdSalir_Click` cierra el formulario de configuración sin pedir confirmación; cambios pendientes se descartan en silencio. | `Form_FormConfiguracion.cls:13-15` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) |
| H-HPS-S-11 | `HaHabidoCambios` exige una comparación manual por cada campo; añadir una propiedad sin agregar el `If <> Nz(...) Then` no marca el cambio, aunque se guarda si otra columna dispara `True`. | `Form_FormConfiguracion.cls:386-457` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) |
| H-HPS-S-12 | `Set m_ObjEntorno.Configuracion = Nothing` solo se invoca al guardar; `cmdSalir` deja la caché de `Configuracion` con la versión anterior en memoria. | `Form_FormConfiguracion.cls:13-15, 72` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) |
| H-HPS-S-13 | `ComandoRegistrar_Click` renueva la caché y emite `Modificado` solo en éxito; un fallo en `Registrar` deja al menú con la versión obsoleta en memoria y sin re-pintado. | `Form_FormConfiguracion.cls:72-78` | [Form_FormConfiguracion](formularios/Form_FormConfiguracion.md) |

### Hallazgos del legado — tanda del alta (H-HPS-S-14 a H-HPS-S-22)

| Id | Hallazgo | Evidencia | Ficha |
|---|---|---|---|
| H-HPS-S-14 | `AvanceCerrar` desde `Form_FormSolicitudAlta` no cierra nada; el selector de tipo queda abierto bajo el destino. | `Form_FormSolicitudAlta.cls:73-84`; `FU.bas:555-582` | [Form_FormSolicitudAlta](formularios/Form_FormSolicitudAlta.md) |
| H-HPS-S-15 | `Form_Open` reactiva `AllowEdits` incondicionalmente, anulando la restricción de perfil técnico. | `Form_FormSolicitudAlta.cls:107, 109-112, 119` | [Form_FormSolicitudAlta](formularios/Form_FormSolicitudAlta.md) |
| H-HPS-S-16 | `DoCmd.Close … acSaveNo` sobre el destino ya abierto descarta trabajo en curso. | `Form_FormSolicitudAlta.cls:69-71` | [Form_FormSolicitudAlta](formularios/Form_FormSolicitudAlta.md) |
| H-HPS-S-17 | Rótulo de diseño de otra app + botón «Ayuda» rotulado «Salir». | `Form_FormSolicitudAlta.form.txt:40, 827-831` | [Form_FormSolicitudAlta](formularios/Form_FormSolicitudAlta.md) |
| H-HPS-S-18 | `ComandoAdjuntarExcel_Click` interpola nombre de fichero y fecha en la SQL. | `SolicitudOperaciones.cls:1516-1525` | [Form_FormAdjuntarExcelSolicitante](formularios/Form_FormAdjuntarExcelSolicitante.md) |
| H-HPS-S-19 | Rótulos heredados: `Caption` de otra app, «Salir»/«Ayuda» invertidos, botones «Extender» y «Editar». | `Form_FormAdjuntarExcelSolicitante.form.txt:40, 289-293, 412, 610-614` | [Form_FormAdjuntarExcelSolicitante](formularios/Form_FormAdjuntarExcelSolicitante.md) |
| H-HPS-S-20 | `EstablecerDatos` decide el estado de botones leyendo controles aún sin pintar. | `Form_FormSolicitudesAltaFechas.cls:1592-1596, 1600, 1612, 1625-1629` | [Form_FormSolicitudesAltaFechas](formularios/Form_FormSolicitudesAltaFechas.md) |
| H-HPS-S-21 | Rótulos residuales: `Caption` de otra app, typo «SOLCITUD», botones `Comando36`/`Comando47` duplicados. | `Form_FormSolicitudesAltaFechas.form.txt:39, 708, 809, 1043` | [Form_FormSolicitudesAltaFechas](formularios/Form_FormSolicitudesAltaFechas.md) |
| H-HPS-S-22 | Los 38 `errores:` repiten «se ha producido el error n»; un handler se identifica con otro nombre. | `Form_FormSolicitudesAltaFechas.cls:703` | [Form_FormSolicitudesAltaFechas](formularios/Form_FormSolicitudesAltaFechas.md) |

### Hallazgos del legado — visor HTML y justificaciones (H-HPS-S-31 a H-HPS-S-41)

| Id | Hallazgo | Evidencia | Ficha |
|---|---|---|---|
| H-HPS-S-31 | `Caption` de diseño «EXPEDIENTES Versión 2024-001 Administrador», de otra aplicación, y ningún handler lo sobrescribe. | `Form_FormWeb.form.txt:35` | [Form_FormWeb](formularios/Form_FormWeb.md) |
| H-HPS-S-32 | Con el centinela interno `Err.Number = 1000`, `m_Error` queda vacío y se muestra igualmente un `MsgBox` crítico sin texto. | `Form_FormWeb.cls:38-44` | [Form_FormWeb](formularios/Form_FormWeb.md) |
| H-HPS-S-33 | Si el HTML a proyectar no existe, el visor se cierra sin comunicar nada; el llamador, que ya pagó la generación, no recibe señal. | `Form_FormWeb.cls:22-26` | [Form_FormWeb](formularios/Form_FormWeb.md) |
| H-HPS-S-34 | `JustificacionRepositorio.Guardar` sólo sabe hacer `INSERT`: editar una justificación reenvía el alta con la clave primaria ya cargada. | `JustificacionRepositorio.bas:35-51` | [Form_frmJustificacion](formularios/Form_frmJustificacion.md) |
| H-HPS-S-35 | `GuardarJustificacion` es un `Sub` que ignora el `Boolean` y el `p_Error` del repositorio: un guardado fallido se comunica como exitoso. | `JustificacionServicio.cls:53` | [Form_frmJustificacion](formularios/Form_frmJustificacion.md) |
| H-HPS-S-36 | El id del alta se reserva con `Max(id)+1` y se recupera con una segunda consulta de `Max` tras el `INSERT`: entre ambas, el id anunciado puede no ser el creado y dos altas simultáneas colisionan. | `JustificacionRepositorio.bas:36, 86-104` | [Form_frmJustificacion](formularios/Form_frmJustificacion.md) |
| H-HPS-S-37 | `Form_Load` no tiene manejador de error: un `OpenArgs` no numérico o un id inexistente producen un error no controlado en vez de un aviso. | `Form_frmJustificacion.cls:17-20, 37-44` | [Form_frmJustificacion](formularios/Form_frmJustificacion.md) |
| H-HPS-S-38 | `IsNull(Me.lstResultados)` comprueba el control y no la selección: sin fila elegida, «Eliminar» pide confirmación y sólo falla después con id `0`. | `Form_frmBuscarJustificacion.cls:156, 172, 238` | [Form_frmBuscarJustificacion](formularios/Form_frmBuscarJustificacion.md) |
| H-HPS-S-39 | La lista se arma con los cinco campos unidos por `;`: un `;` en título o descripción desplaza columnas y rompe el `CBool` que decide si se puede elegir. | `Form_frmBuscarJustificacion.cls:102, 129-133, 299` | [Form_frmBuscarJustificacion](formularios/Form_frmBuscarJustificacion.md) |
| H-HPS-S-40 | Búsqueda y filtro se resuelven recorriendo en memoria la caché completa de justificaciones, que se carga sin `WHERE`: no hay `LIKE` en la base ni paginación. | `Form_frmBuscarJustificacion.cls:98, 106-136` | [Form_frmBuscarJustificacion](formularios/Form_frmBuscarJustificacion.md) |
| H-HPS-S-41 | `txtBuscar_AfterUpdate` actúa sobre el resultado único sin pedir confirmación: una tecla abre el editor o elige la justificación y cierra el buscador. | `Form_frmBuscarJustificacion.cls:51-62` | [Form_frmBuscarJustificacion](formularios/Form_frmBuscarJustificacion.md) |

> Nota: esta tanda arranca en `H-HPS-S-31` porque las fichas-1 y fichas-2 de esta cadena ya publicaron `1..13` y la cadena de `apap3` continúa con su propia numeración. Rango reservado para esta tanda: `H-HPS-S-31` a `H-HPS-S-41`.

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-HPS-1 | ERD de origen: falta la extracción de solo lectura del `.accdb` (workflow `dysflow-erd`, #1144) | `ERD/` del legacy | pendiente del propietario |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/hps-solicitudes/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.
