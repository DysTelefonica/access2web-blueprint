# Form_subfrmDatosPC (grupo de seis subformularios)

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_subfrmDatosPC_{Generales,Propuesta,Impacto,AprobacionSuministrador,DictamenRAC,DecisionFinal}.cls` y `.form.txt` · **PRD:** [README](../README.md)

Seis bloques de la Propuesta de Cambio (`PC`) que se cargan en el `NavigationSubform` de Form_frmDatosPC. Ninguno está enlazado: ningún `.form.txt` declara `RecordSource`, y todos leen y escriben un `DatosPCViewModel` que les pasa el padre (`PoblarControles` / `RecogerValores`). El padre es el que guarda.

**Abreviaturas de cita.** `GEN`, `PRO`, `IMP`, `AS`, `RAC`, `DF` = `Form_subfrmDatosPC_<bloque>.cls`; `PAD` = `forms/Form_frmDatosPC.cls`; `SRV` = `classes/DatosPCServicio.cls`; `REPO` = `modules/DatosPCRepositorio.bas`; `WF` = `classes/WorkflowServicio.cls`; `FPA` = `modules/FormulariosPadreAuxiliares.bas`; `FU` = `modules/FUNCIONES UTILES.bas`; `EXR` = `modules/ExpedienteRepositorio.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_frmDatosPC, pestañas `tabGeneral`… `tabDecisionFinal` | `NavigationButton_Click` → `ConfigurarYPoblarInterfaz` y `PoblarSubformActual`, que invoca `PoblarControles` por `CallByName` | `PAD:134-150, 701-723` |
| Navegación programada del padre | `GetSubformNameFromTabName` + `DoCmd.BrowseTo` | `PAD:100-131`, `FPA:103-150` |

Los subformularios no abren formularios. `DF` lanza, a través del padre, un selector de fichero PDF (`PAD:760-771`).

## Subformularios y controles

| Subformulario | Controles de datos (`.form.txt`) | Botones y handler | Llamadas | Habilitación propia |
|---|---|---|---|---|
| `_Generales` (pestaña «General») | `refContratoInspeccionOficial` y `refSuministrador` con `Tag = "INMUTABLE"` (`.form.txt:621-624, 780-783`); `denominacionContrato`, `suministradorNombreDir`, `objetoContrato` | `cmdGuardar` `GEN:111-124`; `cmdCargarDefaults` `GEN:126-176`; `cmdRestaurarDatos` (oculto por diseño, `.form.txt:1170-1184`) `GEN:178-205` | `SRV.GetDatosPCIniciales` (`SRV:138-231`); `PAD.GuardarDesdeSubform`, `ConfigurarYPoblarInterfaz`, `PoblarSubformActual` | Sí: `AplicarEstadoVisualControl` desactiva siempre los `INMUTABLE` y el resto según `PuedeGuardar_Generales` (`GEN:33-70`) |
| `_Propuesta` («Detalle y Descripción») | `numPlanoEspecificacion`, `descripcionMaterialAfectado`, `descripcionPropuestaCambio` (`.form.txt:662, 691, 782`) | `cmdGuardar` `PRO:27-34` | `PAD.GuardarDesdeSubform` | No: la aplica el padre (`PAD:444, 452-453`) |
| `_Impacto` («Motivos») | 7 casillas `motivo*` + `motivoOtrosDetalle`; grupos `grpIncidenciaCoste`, `grpIncidenciaPlazo` (valores 1-3), `grpClasificacion`, `grpAfeccionMaterial` (1-2); 6 casillas `incidencia*` (`.form.txt:765-1776`) | `cmdGuardar` (desactivado por diseño, `.form.txt:2098-2111`) `IMP:118-130` | `PAD.GuardarDesdeSubform` | No: la aplica el padre (`PAD:445, 452-453`) |
| `_AprobacionSuministrador` | `firmaOficinaTecnicaNombre`, `firmaRepSuministradorNombre` (`.form.txt:642, 737`) | `cmdGuardar` `AS:64-72`; `cmdCargarDefaults` `AS:74-111`; `cmdRestaurarDatos` `AS:113-130` | `SRV.GetDatosAprobacionSuministradorIniciales` (`SRV:101-133`); expone `hayCambios` (`AS:132-141`) que el padre consulta (`PAD:247`) | Sí, según `PuedeGuardar_AprobacionSuministrador` (`AS:32-41`) |
| `_DictamenRAC` («RAC») | `racCodigo`, `racNombre`, `observacionesRAC` (`.form.txt:660, 722, 816`) | `cmdGuardar` `RAC:36-45`; `cmdCargarDatosPredeterminados` `RAC:49-79`; `cmdRestaurarDatos` `RAC:81-99` | `ExpedienteServicio.GetRacAsignadoPorExpediente` → `EXR:79-105` | Sí; `observacionesRAC` queda siempre desactivado (`RAC:18-29`) |
| `_DecisionFinal` («Autoridad Competente») | `NombreAutoridadDiseno`, `obsAprobacionAutoridadDiseno`, `grpDecisionFinal` (1 = APROBADO, 2 = NO APROBADO), `NombreFirmanteFinal`, `obsDecisionFinal` (`.form.txt:653-997`) | `cmdGuardar` `DF:94-121` | `PAD.ProcesarCierreFormalizacion` (`PAD:744-804`) o `PAD.GuardarDesdeSubform` | Sí; todo desactivado si `idEstadoInterno = 8` (`DF:39-60`) |

Los nombres de pestaña salen de `FPA:220-256`. Todos los cuadros de texto declaran `OnChange = "[Procedimiento de evento]"` en el `.form.txt` (por ejemplo `AS .form.txt:644`, `RAC .form.txt:662`) sin handler en su `.cls`.

## Flujos

### Normal: apertura y grabación de un bloque

1. Al pulsar la pestaña, el padre aplica la habilitación genérica (`HabilitarGuardarEnSubform` y `HabilitarControlesSubformulario`, `PAD:452-453`; `FPA:268-273`; `FU:946-980`) y después llama a `PoblarControles`, que vuelca el modelo de vista y, en GEN, AS, RAC y DF, rehace la habilitación (`PAD:136-137, 150`).
2. Con «Grabar», el subformulario llama a `Me.Parent.GuardarDesdeSubform Me.Name` (por ejemplo `PRO:30`). GEN confirma antes la edición en curso con `CommitEditsOnForm` (`GEN:113`; `FU:847-865`). IMP, RAC y DF ejecutan antes `RecogerValores` (`IMP:122`, `RAC:39`, `DF:98`).
3. El padre traduce el nombre a bloque (`PAD:187`; `FPA:160-209`), lee `tbDatosPC`, clona la entidad, vuelve a llamar a `RecogerValores` y copia al objeto solo los campos del bloque (`PAD:191-201, 648-695`).
4. Si la entidad no cambia (`FPA:25-47`), muestra «No se han detectado cambios.» y no guarda, salvo en RAC con `racDecision = "APROBADO"` (`PAD:204-216`). En AS se suma la comprobación `hayCambios` del subformulario (`PAD:246-251`).
5. Llama al servicio del bloque (`PAD:221-255`), muestra «Datos guardados correctamente.» si no hubo avance automático (`PAD:384-386`), recarga el modelo de vista, que recalcula los seis `PuedeGuardar_*` (`PAD:573-578`), y repuebla el subformulario (`PAD:389-396`).

### Alternativos

- **Generales con avance de estado.** `GuardarDatosGenerales` guarda y, si la solicitud está en estado 1, la pasa a 2 dentro de la misma transacción (`SRV:259-276`). Después, en estado 2, el padre pregunta «¿Asignar tarea al Técnico ahora?» y, si se acepta, pasa a 3 (`PAD:265-281`).
- **Propuesta e Impacto en Desarrollo Técnico (estado 3).** El padre comprueba si están completos los motivos y el detalle, navega a la pestaña que falta o pregunta si se envía a Calidad, lo que pasa la solicitud a 4 (`PAD:287-332, 339-378`).
- **Cargar datos predeterminados.** GEN propone `CodExpLargo`, código de solicitud, título y objeto del expediente, y un suministrador compuesto: subcontratistas si existen y, si no, contratistas principales; con más de uno, «UNION TEMPORAL DE EMPRESAS: A y B»; la dirección sale del entorno o del primer suministrador (`SRV:154-219`). AS propone el responsable técnico del expediente y el jefe de calidad del entorno con sufijo de rol (`SRV:118-123, 871-883`). RAC propone el RAC asignado al expediente o avisa «Este expediente no tiene un RAC asignado…» (`RAC:61-67`). En los tres casos se guarda el estado previo, se desactiva el botón de carga y se muestra «Restaurar» (`GEN:138-165`, `AS:85-101`, `RAC:59-71`). Nada se persiste hasta «Grabar».
- **Restaurar.** Devuelve a los controles los valores guardados antes de la carga (`GEN:182-200`, `AS:117-123`, `RAC:84-93`).
- **AS en Modificación (estado 4).** Tras guardar, recarga Form_frmGestionSolicitud si está abierto (`PAD:258-262`).
- **Decisión final APROBADO.** `DF:107-110` delega en `ProcesarCierreFormalizacion`: confirma, exige elegir un PDF (`PAD:753-771`) y `EjecutarCierreFormalizacion` guarda la decisión, sustituye los adjuntos de cierre anteriores, registra el nuevo y pasa la solicitud a 8, todo en una transacción (`WF:3347-3385`).
- **Decisión final NO APROBADO.** Solo guarda la decisión, sin transición (`DF:112-116`; `PAD:254`).

### Errores

- Validaciones de servicio con error 513, mostradas por `ManejadorErroresFormulario` (`PAD:410-418`; `FU:588-613`): los cinco campos de Generales (`SRV:253-257`); material afectado y propuesta (`SRV:303-308`); reglas de Impacto (`SRV:377-419`); las dos firmas (`SRV:499-506`); código RAC (`SRV:526-530`); decisión final (`SRV:565`).
- DF valida antes en pantalla: «El campo 'Decisión Final' es obligatorio.» (`DF:101-104`).
- Sin registro en `tbDatosPC`, Propuesta, Impacto y AS fallan con «No existe registro en tbDatosPC para esta solicitud. Debes guardar primero el bloque 'Datos Generales'.» (`REPO:187-191, 240-245, 306-311`).
- En RAC, un error revierte la transacción abierta por el padre (`PAD:221-232, 412-414`). En la carga de valores predeterminados, el error se muestra y se libera el bloqueo visual (`GEN:173-175`, `AS:108-110`, `RAC:76-78`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Quién edita cada bloque: nadie en estado 9; en 8, solo Decisión final (Calidad o Administrador); en 7, solo Aprobación suministrador y Decisión final; antes de 7: Generales para Calidad o Administrador; Propuesta e Impacto para Técnico en 3 y para Calidad o Administrador en 4 o más; Aprobación suministrador y RAC para Calidad o Administrador desde 4; Decisión final, nadie | `WF.PermiteEdicion_PC` | `WF:1419-1445, 1599-1670`; estados en `modules/modEnumeradores.bas:36-46` |
| Impacto: al menos un motivo; con «Otros», detalle obligatorio; detalle sin «Otros», error; coste y plazo ∈ {Aumentará, Disminuirá, No Variará}; clasificación MAYOR o MENOR; material «ya entregado» o «por entregar» | `SRV.GuardarImpacto` | `SRV:377-419` |
| Correspondencia entre botón de opción y texto guardado | `IMP`, `DF` | `IMP:21-56, 78-109`; `DF:26-30, 76-80` |
| Código RAC obligatorio salvo `racDecision = "RECHAZADO"`; el subformulario no edita `racDecision` | `SRV`, `RAC` | `SRV:526-530`; `RAC:102`; se asigna en `forms/Form_frmGestionSolicitud.cls:2185` y `SRV:633` |
| Con `Tag = "INMUTABLE"`, el control nunca se habilita | `GEN`, `FU` | `GEN:35-38`; `FU:954-956` |
| Trazabilidad de rechazo: si hay rechazo activo, Propuesta e Impacto anotan cada campo cambiado en el JSON de rechazo | `SRV` | `SRV:311-345, 421-460` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del bloque | `tbDatosPC` | DAO (por servicio) | No | `PAD:191`; `SRV:664` |
| Generales: alta o modificación de toda la fila; id = `Max(idDatosPC) + 1` | `tbDatosPC` | DAO `AddNew`/`Edit`, SQL dinámico | Sí, con la transición 1→2 | `REPO:10-50, 75-86`; `SRV:259-276` |
| Propuesta | `tbDatosPC` (3 campos) | DAO con `QueryDef` parametrizada | Solo en estado 5 | `REPO:176-202`; `SRV:322-333` |
| Impacto | `tbDatosPC` (20 campos) | DAO, SQL dinámico | Solo en estado 5 | `REPO:236-271`; `SRV:463-474` |
| Aprobación suministrador | `tbDatosPC` (2 firmas) | DAO, SQL dinámico | No | `REPO:302-318` |
| Dictamen RAC | `tbDatosPC` (`racCodigo`, `observacionesRAC`, `racNombre`, `racDecision`, `racRechazoMotivos`) | DAO, SQL dinámico | Sí, abierta por el padre | `REPO:353-371`; `PAD:221-232` |
| Decisión final | `tbDatosPC` (5 campos) | DAO con `QueryDef` parametrizada | En APROBADO, junto con los adjuntos y la transición a 8 | `REPO:394-412`; `WF:3352-3385` |
| Estado para decidir la transacción | `tbSolicitudes` | SQL dinámico | — | `SRV:841-869` |
| RAC predeterminado | `TbExpedientesRACS` ⋈ `TbRACS` en la base de expedientes | SQL dinámico con `FIRST()` | — | `EXR:86-91` |

## Coexistencia

Pendiente de `coexistencia.md`. Lo que ya se ve: el agregado `FIRST()` de Jet en `EXR:87`, que no existe en PostgreSQL; el identificador `Max + 1` (`REPO:80-86`), que no es seguro con concurrencia y debe sustituirse por una secuencia; las transacciones DAO con `ws.OpenDatabase` sobre la base actual (`PAD:222-224`; `SRV:260-262`), que deben comprobarse con tablas vinculadas; dos orígenes de datos distintos (`getdb` y `getdbExpedientes`, `FU:84, 156`); y booleanos de Access en casillas (`REPO:249-265`). Debe conservarse que cada bloque actualice solo sus columnas (`PAD:648-695`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- IMP usa `m_ViewModel` (`IMP:122`) sin declararlo, con `Option Explicit` (`IMP:3`), y no existe una variable global con ese nombre en `modules` ni en `classes`. Debería dar «Variable no definida» al compilar.
- En `IMP:103`, si no hay clasificación se vacía `vm_incidenciaPlazo` en lugar de `vm_impactoClasificacion`. El servicio rechaza entonces por el plazo (`SRV:402-406`) antes de llegar al mensaje de la clasificación (`SRV:409-413`).
- `ImportarEntidad_aVM` asigna `NombreAutoridadDiseno` a `vm_decisionFinalNombreFirmante` (`PAD:642`), que se sobrescribe en `PAD:645`. `vm_decisionAutoridadDisenioNombre` no se carga nunca: el control aparecería vacío (`DF:22`) y el siguiente guardado escribiría el vacío (`PAD:690`; `REPO:408`).
- En APROBADO, `ProcesarCierreFormalizacion` copia cuatro campos (`PAD:781-784`) pero no `obsAprobacionAutoridadDiseno`: la observación tecleada no se guarda (`REPO:407`).
- `observacionesRAC` está siempre desactivado (`RAC:23`), pero se recoge y se persiste (`RAC:107`; `REPO:364`).
- `AS.GuardarAprobacionSuministrador` no tiene `Exit Sub` antes de `Errores:` y siempre llama al manejador (`AS:143-149`). No tiene llamadas en `src`.
- DF compara con el literal 8 (`DF:40`) en lugar de `estadoAprobada` (`modEnumeradores.bas:44`).
- El borrado de adjuntos y el registro del PDF están dentro de la transacción (`WF:3372-3380`), pero la copia del fichero, no. Queda por verificar si un fallo posterior deja un fichero huérfano.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un formulario por bloque con validación en cliente y en servidor; la regla de permisos de `WF:1599-1670` como única fuente; la subida del PDF de cierre y la transición en una sola operación del servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Cada bloque se edita solo con el rol y el estado de la tabla de permisos; si no, los campos y «Grabar» aparecen desactivados (`WF:1599-1670`).
2. Las dos referencias de Generales nunca son editables (`GEN:35-38`).
3. «Cargar datos predeterminados» propone los valores de `SRV:154-219`, `SRV:118-123` o `EXR:86-91` sin persistirlos, y «Restaurar» recupera los anteriores.
4. Guardar sin cambios muestra «No se han detectado cambios.» y no escribe (`PAD:204-216`).
5. Cada mensaje de validación de `SRV:253-565` aparece con su dato inválido y no se escribe nada.
6. Guardar Generales en estado 1 pasa la solicitud a 2 en la misma operación (`SRV:269-273`).
7. APROBADO exige un PDF y deja la decisión, el adjunto y el estado 8 de forma atómica; NO APROBADO solo guarda (`WF:3347-3385`; `DF:107-116`).
8. Con la solicitud en estado 8, la decisión final es de solo lectura (`DF:39-60`).

## Pendientes

- Compilar el binario para confirmar si IMP compila y si Access ignora los `OnChange` declarados sin handler.
- Si NO APROBADO debe llevar a Rechazada (estado 9): el código no lo hace.
- Columnas que escribe `RepositorioComun.RellenarRecordsetDesdeObjeto` en Generales (`REPO:32, 37`).
- Contenido de `EjecutarTransicion` (`WF:1211`) y de los adjuntos, que se documenta en las fichas de flujo de trabajo y de adjuntos.
