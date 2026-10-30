# Form_frmDatosCDCASUB

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDatosCDCASUB.cls` (779 líneas) y `.form.txt` · **PRD:** [README](../README.md) · **Variante principal:** [Form_frmDatosCDCA](Form_frmDatosCDCA.md)

Formulario contenedor de los datos técnicos de una concesión o desviación de **subsuministrador** (tipo `CD_CA_SUB`). Tiene la misma estructura que `Form_frmDatosCDCA`: seis pestañas, subformularios `subfrmDatosCDCASUB_*`, ViewModel y guardado por bloque. Esta ficha describe lo que difiere; lo común se remite a la ficha de la variante principal. Es **no enlazado** (sin `RecordSource` en el `.form.txt`).

Abreviaturas: sin prefijo, `Form_frmDatosCDCASUB.cls`; `WF`, `classes/WorkflowServicio.cls`; `SSrv`, `classes/DatosCDCASUBServicio.cls`; `SRepo`, `modules/DatosCDCASUBRepositorio.bas`; `GS`, `forms/Form_frmGestionSolicitud.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Gestión de solicitud, botón «Ver datos» | Si `tipoSolicitud = "CD_CA_SUB"`, abre `frmDatosCDCASUB` con el id en `OpenArgs` | `GS:305, 335` |
| Gestión de solicitud | Escucha `SolicitudHaCambiado` | `GS:37, 346, 2335-2357` |
| Subformularios `subfrmDatosCDCASUB_*` | `Me.Parent.GuardarDesdeSubform`; con Decisión final `APROBADO`, `Me.Parent.ProcesarCierreFormalizacion` | `subfrmDatosCDCASUB_DecisionFinal.cls:68-72`; `subfrmDatosCDCASUB_Generales.cls:118` |

Devuelve el foco a `frmGestionSolicitud` igual que la variante principal (`:149-168`).

## Controles y eventos

Mismos controles, nombres y textos que la variante principal (`.form.txt:650-1301`): pestañas `tabGeneral` … `tabDecisionFinal` con destino `subfrmDatosCDCASUB_*` (`.form.txt:704-1004`), `NavigationSubform` (`.form.txt:1091`), `cmdCerrar` (`:144-147`), `cmdVolverResumen` (`:149-168`), `txtEstado`, `lblBloqueo` y `lblBannerRechazo`. Diferencias:

| Elemento | Diferencia | Evidencia |
|---|---|---|
| `Load` | Sin id en `OpenArgs` no da error: deja el formulario vacío a la espera de `CargarDatos` | `:29-35` |
| `CargarDatos` (público) | Recarga, configura la interfaz y navega a la pestaña del estado | `:47-70` |
| Título | «Concesión/Desviación SUB: …» | `:84` |
| Errores | `ErrorPresenter.MostrarError` en lugar de `ManejadorErroresFormulario` | `:41, 69, 129, 481` |

## Flujos

Apertura, guardado y cierre siguen los pasos de la variante principal (`:19-70`, `:246-482`, `:483-537`). Cambia lo siguiente:

1. **Estado y tiempo de ciclo.** El nombre del estado sale de la caché `m_ObjEntorno.estados`, con «Estado n» si falta (`:659-663`). Para Aprobada o Rechazada, el ciclo termina en la última transición del historial (`:666-675`).
2. **Control de cambios.** Aprobación suministrador usa solo `hayCambios()` del subformulario; el resto compara entidades (`:287-305`).
3. **Dictamen RAC.** El formulario no abre transacción; el servicio la abre solo si la solicitud está en Validación (`:308-310`; `SSrv:337-352`).
4. **Traza de edición.** Si no hubo transición automática, al guardar Propuesta, Impacto o Generales escribe un registro con tabla `tbDatosCDCASUB`, valores fijos «...» y «Actualizado» y tipo `EDICION_TECNICA` (`:333-346`).
5. **Estado tras guardar.** Relee la solicitud de la base antes de decidir avances (`:349`) y actualiza el estado en memoria tras cada transición (`:360, 440`).
6. **Reintento de navegación.** Si `BrowseTo` falla, intenta reabrir el propio formulario y repetir (`:220-227`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Generales obligatorios: referencia del subsuministrador, suministrador principal, subsuministrador y «requiere modificación»; la referencia del suministrador ya no lo es | Servicio | `SSrv:141-145` |
| «Requiere modificación» se normaliza: booleano, número distinto de 0 o nulo | Formulario | `:558-568` |
| Generales en Preregistro transita a Registro en la misma transacción | Servicio | `SSrv:147-160` |
| Propuesta obligatoria: identificación del material, causa de la no conformidad y descripción del impacto | Servicio | `SSrv:180-182` |
| Impacto obligatorio: coste, clasificación y autoridad de diseño del subsuministrador (`esSubSuministradorAD`) no vacíos | Servicio | `SSrv:248-250` |
| Aprobación: solo Ingeniería y Calidad obligatorias; sin indicadores «aplica» | Servicio | `SSrv:317-325`; `:322` |
| RAC: código obligatorio salvo `RECHAZADO`; añade observaciones y nombre del delegador | Servicio y formulario | `SSrv:334-336`; `:606-608` |
| Parte técnica completa: material, causa, descripción, clasificación y autoridad de diseño; **no** comprueba el coste | Servicio | `SSrv:401-430` |
| Motivos completos: sí comprueba el coste | Servicio | `SSrv:452-475` |
| Edición por rol y estado: mismas reglas que `CD_CA`, en función propia | Workflow | `WF:1557-1598` |
| Precondiciones de transición propias del tipo | Workflow | `WF:926-953` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura | `tbDatosCDCASUB` | SQL parametrizado `[p_ID]` | No | `SRepo:62-79` |
| Generales (alta o edición) | `tbDatosCDCASUB` | DAO `AddNew/Edit`, id = `Max+1` | Sí (servicio) | `SRepo:6-27, 28-60` |
| Propuesta, Impacto, Aprobación, RAC, Decisión | `tbDatosCDCASUB` | DAO `Edit` con SQL concatenado | Según bloque | `SRepo:127-317` |
| Traza de edición | log de cambios | `LogServicio.RegistrarCambio` | No | `:344`; `classes/LogServicio.cls:18-38` |
| Cambio de estado, cierre y adjunto | como la variante principal | — | Sí | `WF:1211-1409, 3334-3397` |

## Coexistencia

Pendiente de `../coexistencia.md`. Mismos puntos que la variante principal: transacciones DAO con conexión propia (`SSrv:147-149, 340-342`), id por `Max+1` (`SRepo:11`) y SQL concatenado (`SRepo:34`). Además, la traza escribe en la misma ejecución que el guardado pero fuera de su transacción (`:344`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba y en la ficha de la variante principal.

**Posibles defectos (sin corregir, a verificar):**

- `RecargarViewModel` crea el `CondorError` pero no lo lanza (`:693-695`): un fallo de lectura deja el ViewModel anterior o vacío sin aviso, y `CargarDatos` falla después en `:58`.
- No se ha encontrado ningún llamador externo de `CargarDatos`; abrir sin `OpenArgs` deja un formulario vacío (`:32-35`).
- Con rechazo previo y Motivos incompletos por el coste, el escenario 1 no actúa (`:383`) y `EsParteTecnicaCompleta`, que no mira el coste (`SSrv:417-421`), puede ofrecer el envío a Calidad (`:420-442`).
- La traza dice ser «vital para desbloqueo técnico» (`:333`), pero `VerificarCambiosTrasRechazo` ya no consulta el log (`WF:3425-3451`); registra valores fijos, no los reales (`:344`).
- El reintento de `:223-227` no se ejecuta: `On Error GoTo 0` (`:218`) propaga el error de `BrowseTo` antes de evaluarlo.
- Tras enviar a Calidad o navegar se sale sin `RaiseEvent` ni recarga (`:386, 393, 405, 412, 449`), como en la variante principal.
- `mostrarBtnTecnica` se calcula y no se usa (`:112-116`); el comentario de `:701` nombra `tbDatosCDCA_SUB` y el repositorio usa `tbDatosCDCASUB` (`SRepo:34`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: una sola pantalla parametrizada por tipo (`CD_CA` / `CD_CA_SUB`) con los campos y validaciones de cada variante.

## Criterios de aceptación

Propuestos, no aprobados:

1. Generales exige los tres datos de subsuministro y «requiere modificación»; no exige la referencia del suministrador.
2. Aprobación suministrador exige solo Ingeniería y Calidad.
3. El RAC guarda también observaciones y nombre del delegador.
4. Un fallo al cargar la solicitud muestra un error y no presenta datos de otra solicitud.
5. Los criterios 2-9 de la variante principal se cumplen con `tbDatosCDCASUB` y el tipo `CD_CA_SUB`.

## Pendientes

- Si la ausencia de comprobación del coste en `EsParteTecnicaCompleta` es intencionada (`SSrv:416-421`).
- Precondiciones concretas `CD_CA_SUB` (`WF:926-953`).
- Deriva del código: el `CodeBehind` del `.form.txt` (`.form.txt:1319`) usa helpers locales y el `.cls`, los de `FormulariosPadreAuxiliares`.
- Contenido de `GetDatosGeneralesIniciales` (`SSrv:24-113`).
