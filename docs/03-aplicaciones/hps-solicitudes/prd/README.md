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
| Form_Form0BDOpciones | src/forms/Form_Form0BDOpciones.cls | 407 | 33 | 13 | igual | pendiente | — |
| Form_FormAdjuntaTraspasoONS | src/forms/Form_FormAdjuntaTraspasoONS.cls | 234 | 28 | 7 | igual | pendiente | — |
| pendiente | — ||| pendiente | — | | pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |A| pendiente | — |d| pendiente | — |j| pendiente | — |u| pendiente | — |n| pendiente | — |t| pendiente | — |a| pendiente | — |r| pendiente | — |E| pendiente | — |x| pendiente | — |c| pendiente | — |e| pendiente | — |l| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |a| pendiente | — |n| pendiente | — |t| pendiente | — |e| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |s| pendiente | — |r| pendiente | — |c| pendiente | — |/| pendiente | — |f| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |s| pendiente | — |/| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |A| pendiente | — |d| pendiente | — |j| pendiente | — |u| pendiente | — |n| pendiente | — |t| pendiente | — |a| pendiente | — |r| pendiente | — |E| pendiente | — |x| pendiente | — |c| pendiente | — |e| pendiente | — |l| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |a| pendiente | — |n| pendiente | — |t| pendiente | — |e| pendiente | — |.| pendiente | — |c| pendiente | — |l| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |2| pendiente | — |1| pendiente | — |5| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |2| pendiente | — |8| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |7| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |a| pendiente | — |l| pendiente | — | | pendiente | — || pendiente | — ||| pendiente | — |
| Form_FormAltaDatosSolicitante | src/forms/Form_FormAltaDatosSolicitante.cls | 452 | 63 | 15 | igual | pendiente | — |
| Form_FormCambioTipo | src/forms/Form_FormCambioTipo.cls | 107 | 26 | 4 | igual | pendiente | — |
| pendiente | — ||| pendiente | — | | pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |C| pendiente | — |o| pendiente | — |n| pendiente | — |f| pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |r| pendiente | — |a| pendiente | — |c| pendiente | — |i| pendiente | — |o| pendiente | — |n| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |s| pendiente | — |r| pendiente | — |c| pendiente | — |/| pendiente | — |f| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |s| pendiente | — |/| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |C| pendiente | — |o| pendiente | — |n| pendiente | — |f| pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |r| pendiente | — |a| pendiente | — |c| pendiente | — |i| pendiente | — |o| pendiente | — |n| pendiente | — |.| pendiente | — |c| pendiente | — |l| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |4| pendiente | — |1| pendiente | — |8| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |5| pendiente | — |0| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |1| pendiente | — |1| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |a| pendiente | — |l| pendiente | — | | pendiente | — || pendiente | — ||| pendiente | — |
| Form_FormCorreo | src/forms/Form_FormCorreo.cls | 403 | 36 | 10 | igual | analizado | [formularios/Form_FormCorreo.md](formularios/Form_FormCorreo.md) |
| Form_FormCorreoPruebas | src/forms/Form_FormCorreoPruebas.cls | 259 | 38 | 8 | igual | analizado | [formularios/Form_FormCorreoPruebas.md](formularios/Form_FormCorreoPruebas.md) |
| Form_FormCorreosGestion | src/forms/Form_FormCorreosGestion.cls | 541 | 34 | 19 | igual | analizado | [formularios/Form_FormCorreosGestion.md](formularios/Form_FormCorreosGestion.md) |
| Form_FormEntidades | src/forms/Form_FormEntidades.cls | 170 | 17 | 7 | igual | pendiente | — |
| Form_FormEnvioCorreo | src/forms/Form_FormEnvioCorreo.cls | 428 | 36 | 10 | igual | analizado | [formularios/Form_FormEnvioCorreo.md](formularios/Form_FormEnvioCorreo.md) |
| Form_FormExpedienteDetalle | src/forms/Form_FormExpedienteDetalle.cls | 79 | 32 | 3 | igual | pendiente | — |
| Form_FormExpedientesBusqueda | src/forms/Form_FormExpedientesBusqueda.cls | 352 | 27 | 10 | igual | pendiente | — |
| Form_FormMotivoHPS | src/forms/Form_FormMotivoHPS.cls | 124 | 23 | 4 | igual | pendiente | — |
| Form_FormMotivosHPS | src/forms/Form_FormMotivosHPS.cls | 323 | 18 | 10 | igual | pendiente | — |
| Form_FormPlantillasHTML | src/forms/Form_FormPlantillasHTML.cls | 181 | 17 | 7 | igual | pendiente | — |
| Form_FormResponsable | src/forms/Form_FormResponsable.cls | 214 | 28 | 6 | igual | pendiente | — |
| Form_FormResponsablesGestion | src/forms/Form_FormResponsablesGestion.cls | 392 | 28 | 13 | igual | pendiente | — |
| pendiente | — ||| pendiente | — | | pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |u| pendiente | — |d| pendiente | — |A| pendiente | — |l| pendiente | — |t| pendiente | — |a| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |s| pendiente | — |r| pendiente | — |c| pendiente | — |/| pendiente | — |f| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |s| pendiente | — |/| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |u| pendiente | — |d| pendiente | — |A| pendiente | — |l| pendiente | — |t| pendiente | — |a| pendiente | — |.| pendiente | — |c| pendiente | — |l| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |1| pendiente | — |3| pendiente | — |0| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |2| pendiente | — |4| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |4| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |a| pendiente | — |l| pendiente | — | | pendiente | — || pendiente | — ||| pendiente | — |
| Form_FormSolicitudAltaTraspaso | src/forms/Form_FormSolicitudAltaTraspaso.cls | 142 | 25 | 5 | igual | pendiente | — |
| Form_FormSolicitudRenovacion | src/forms/Form_FormSolicitudRenovacion.cls | 151 | 25 | 5 | igual | pendiente | — |
| Form_FormSolicitudRenovacionDatos | src/forms/Form_FormSolicitudRenovacionDatos.cls | 957 | 64 | 19 | igual | pendiente | — |
| Form_FormSolicitudRenovacionFechas | src/forms/Form_FormSolicitudRenovacionFechas.cls | 1652 | 82 | 40 | igual | pendiente | — |
| pendiente | — ||| pendiente | — | | pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |u| pendiente | — |d| pendiente | — |e| pendiente | — |s| pendiente | — |A| pendiente | — |l| pendiente | — |t| pendiente | — |a| pendiente | — |F| pendiente | — |e| pendiente | — |c| pendiente | — |h| pendiente | — |a| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |s| pendiente | — |r| pendiente | — |c| pendiente | — |/| pendiente | — |f| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |s| pendiente | — |/| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |o| pendiente | — |l| pendiente | — |i| pendiente | — |c| pendiente | — |i| pendiente | — |t| pendiente | — |u| pendiente | — |d| pendiente | — |e| pendiente | — |s| pendiente | — |A| pendiente | — |l| pendiente | — |t| pendiente | — |a| pendiente | — |F| pendiente | — |e| pendiente | — |c| pendiente | — |h| pendiente | — |a| pendiente | — |s| pendiente | — |.| pendiente | — |c| pendiente | — |l| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |1| pendiente | — |6| pendiente | — |6| pendiente | — |1| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |8| pendiente | — |2| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |4| pendiente | — |0| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |a| pendiente | — |l| pendiente | — | | pendiente | — || pendiente | — ||| pendiente | — |
| Form_FormSolicitudesAltaRenovacionFechas | src/forms/Form_FormSolicitudesAltaRenovacionFechas.cls | 751 | 63 | 18 | igual | pendiente | — |
| Form_FormSolicitudesAltaTraspasoDatos | src/forms/Form_FormSolicitudesAltaTraspasoDatos.cls | 953 | 66 | 23 | igual | pendiente | — |
| Form_FormSolicitudesGestion | src/forms/Form_FormSolicitudesGestion.cls | 912 | 48 | 23 | igual | pendiente | — |
| Form_FormSolicitudesTraspasoFechas | src/forms/Form_FormSolicitudesTraspasoFechas.cls | 1166 | 56 | 29 | igual | pendiente | — |
| Form_FormTareasTramitadorPendientes | src/forms/Form_FormTareasTramitadorPendientes.cls | 652 | 29 | 14 | igual | pendiente | — |
| Form_FormUsuariosHPSBusqueda | src/forms/Form_FormUsuariosHPSBusqueda.cls | 295 | 26 | 9 | igual | pendiente | — |
| Form_FormWeb | src/forms/Form_FormWeb.cls | 56 | 16 | 2 | igual | pendiente | — |
| Form_frmBuscarJustificacion | src/forms/Form_frmBuscarJustificacion.cls | 333 | 29 | 17 | igual | pendiente | — |
| Form_frmJustificacion | src/forms/Form_frmJustificacion.cls | 139 | 28 | 6 | igual | pendiente | — |
| pendiente | — ||| pendiente | — | | pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |f| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |p| pendiente | — |l| pendiente | — |a| pendiente | — |s| pendiente | — |h| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |s| pendiente | — |r| pendiente | — |c| pendiente | — |/| pendiente | — |f| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |s| pendiente | — |/| pendiente | — |F| pendiente | — |o| pendiente | — |r| pendiente | — |m| pendiente | — |_| pendiente | — |f| pendiente | — |r| pendiente | — |m| pendiente | — |S| pendiente | — |p| pendiente | — |l| pendiente | — |a| pendiente | — |s| pendiente | — |h| pendiente | — |.| pendiente | — |c| pendiente | — |l| pendiente | — |s| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |6| pendiente | — |3| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |1| pendiente | — |5| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |2| pendiente | — | | pendiente | — ||| pendiente | — | | pendiente | — |i| pendiente | — |g| pendiente | — |u| pendiente | — |a| pendiente | — |l| pendiente | — | | pendiente | — || pendiente | — ||| pendiente | — |

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

Totales: 14 analizados · 0 no aplicables · 60 pendientes.



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

### Hallazgos del legado — tanda de correos (H-HPS-S-31 a H-HPS-S-58)

| Id | Hallazgo | Evidencia | Ficha |
|---|---|---|---|
| H-HPS-S-31 | `ComandoReenviar_Click` clona `cadenaRecursos` y `URLAdjunto` antiguos: si la plantilla o el adjunto cambiaron, el clon usa los valores obsoletos. | `Form_FormCorreo.cls:151-174` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-32 | Los `Comando*_Click` repiten el patrón de error copy-paste: literal «Al ComandoAyuda_Click se ha devuelto el error» reutilizado en todos los handlers. | `Form_FormCorreo.cls:42, 132-137` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-33 | `ComandoReenviar_Click` no recalcula `cadenaRecursos` desde la solicitud origen: la cadena queda con los `||parametros||` del correo original aunque la solicitud haya cambiado. | `Form_FormCorreo.cls:165` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-34 | `HaHabidoCambios` solo compara 5 campos y devuelve `True` con la primera diferencia: ediciones que solo afectan a campos no comparados pasan inadvertidas. | `Form_FormCorreo.cls:187-235` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-35 | `BloquearControles True` se ejecuta con `On Error Resume Next`: controles no presentes en este `.form.txt` (p. ej. `DestinatariosConCopiaOculta`) no se bloquean por omisión silenciosa. | `Form_FormCorreo.cls:305-318` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-36 | `Me.OpenArgs` se valida como string literal `= "MODO_GESTION"`: el modo se transmite sin validar origen, cualquier llamante puede abrir con ese `OpenArgs` y desactivar la edición. | `Form_FormCorreo.cls:60-65, 278-285` | [Form_FormCorreo](formularios/Form_FormCorreo.md) |
| H-HPS-S-37 | El combo de solicitudes concatena `Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')'` sin protección: un solicitante sin `Apellido1` muestra el ID pegado al nombre con espacio doble. | `Form_FormCorreoPruebas.cls:41-42` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-38 | `cmdVerHTML_Click` aborta con `MsgBox` si no hay `m_SolicitudOrigen`, pero `cmdEnviarPrueba_Click` aborta con un `MsgBox` distinto y sin la misma protección homogénea. | `Form_FormCorreoPruebas.cls:161-165, 195-200` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-39 | `cmdEnviarPrueba_Click` no valida que `p_FechaOrden` no esté en el pasado: una fecha anterior se persiste como programada. | `Form_FormCorreoPruebas.cls:228-232` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-40 | `Form_Load` no usa `On Error GoTo errores:`: un fallo en el `OpenRecordset` se maneja con `MsgBox` y deja `rcd` sin cerrar. | `Form_FormCorreoPruebas.cls:23-79` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-41 | El `Caption` del formulario es literal `"CONSOLA DE ENVÍO DE CORREOS DE PRUEBA"`, sin el sufijo de versión de `m_TituloFormulario` que sí se usa en `Form_FormCorreo` y `Form_FormEnvioCorreo`. | `Form_FormCorreoPruebas.cls:25` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-42 | `CONST_IDSOLICITUD_PRUEBA` se referencia pero no se define en este `.cls`. A verificar su valor y ubicación. | `Form_FormCorreoPruebas.cls:241` | [Form_FormCorreoPruebas](formularios/Form_FormCorreoPruebas.md) |
| H-HPS-S-43 | `ComandoLimpiarSolicitud` y `ComandoLimpiarEstado` solo actúan si el valor es distinto del destino: el usuario ve un botón inerte sin feedback. | `Form_FormCorreosGestion.cls:67-83` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-44 | `ComandoEliminar_Click` permite eliminar correos «Programados» si la columna `FechaEnvio` aún está vacía: una entrega en cola puede llegar a hacerse entre el click y el `DELETE`. | `Form_FormCorreosGestion.cls:325-350` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-45 | `ComandoInformeControl_Click` registra log con `m_ObjUsuarioConectado.UsuarioRed` literal: si el usuario de red está en blanco o coincide con el de Windows, la traza no diferencia sesiones. | `Form_FormCorreosGestion.cls:42` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-46 | `m_Form_Alta` y `m_Form_Editado` invocan `Filtrar` sin argumento: el `ByRef p_Error As String` queda sin definir; a verificar el comportamiento real. | `Form_FormCorreosGestion.cls:397-403` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-47 | El formulario tiene `Estado_BeforeUpdate` y `Estado_AfterUpdate` que llaman ambos a `Filtrar`: un futuro cambio en `Filtrar` que modifique `Estado` produciría bucle. | `Form_FormCorreosGestion.cls:75-78, 372-375` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-48 | `ComandoInformeControl_Click` no comprueba permisos: cualquier usuario con acceso al menú puede generar el informe. | `Form_FormCorreosGestion.cls:23-58` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-49 | `ComandoLimpiarSolicitud` y `ComandoLimpiarEstado` no se deshabilitan cuando ya están en su valor por defecto: el usuario ve un botón inerte. | `Form_FormCorreosGestion.cls:80-83` | [Form_FormCorreosGestion](formularios/Form_FormCorreosGestion.md) |
| H-HPS-S-50 | `ComandoEnviar_Click` refresca el padre según `TIPO`: para `Renovación` cae en el bloque «resto» que invoca `Form_FormSolicitudesAltaRenovacionFechas.EstablecerDatos`, sin cubrir otras pantallas de renovación abiertas en paralelo. | `Form_FormEnvioCorreo.cls:113-119` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-51 | `Registrar` aborta con «Acción no codificada» ante un `m_TipoEnvioCorreo` desconocido: el usuario pierde el trabajo; debería mantener el formulario abierto. | `Form_FormEnvioCorreo.cls:275-277` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-52 | `HabilitarEdicionFecha` desenlaza el control cuando el `ControlSource` parece calculado: con `Form_Current` re-ejecutándose en cada cambio de registro, el desenlace ocurre múltiples veces. | `Form_FormEnvioCorreo.cls:424-441` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-53 | `ComandoEnviar_Click` usa `On Error Resume Next` solo en el bloque de refresco del padre: los errores al refrescar quedan silenciados. | `Form_FormEnvioCorreo.cls:108-119` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-54 | `Registrar` no actualiza `m_ObjSolicitudActiva.SolicitudFechas.{columna}` antes de volver: la versión en memoria queda obsoleta hasta el refresco final. | `Form_FormEnvioCorreo.cls:244-292` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-55 | `Me.Caption = m_TituloFormulario` y la asignación posterior a `m_TituloFormulario` con `Me.lblTitulo.Caption` compiten en `Form_Open`: a verificar el orden de pintado. | `Form_FormEnvioCorreo.cls:174-181, 375` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-56 | `chkActualizarFecha.Enabled = False` cuando el hito no existe: el usuario no puede desactivar la actualización, pero el `ComandoGrabarFecha` tiene su propio caption sin restricción explícita. | `Form_FormEnvioCorreo.cls:220-222, 310-352` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-57 | `ComandoGrabarFecha_Click` no refresca el padre tras grabar: la pantalla `Form_FormSolicitudesAltaFechas` abierta detrás no se actualiza con la nueva fecha. | `Form_FormEnvioCorreo.cls:310-340` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |
| H-HPS-S-58 | `ValidarFechaEdicion` rechaza fechas futuras pero acepta cualquier fecha anterior: una fecha anterior a hoy se persiste como reenvío válido. | `Form_FormEnvioCorreo.cls:407-422` | [Form_FormEnvioCorreo](formularios/Form_FormEnvioCorreo.md) |

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-HPS-1 | ERD de origen: falta la extracción de solo lectura del `.accdb` (workflow `dysflow-erd`, #1144) | `ERD/` del legacy | pendiente del propietario |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/hps-solicitudes/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.
