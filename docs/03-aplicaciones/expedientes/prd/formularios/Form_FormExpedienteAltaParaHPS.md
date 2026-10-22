# Form_FormExpedienteAltaParaHPS

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteAltaParaHPS.cls` (453 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario simplificado de alta de un expediente del tipo «Expediente Sólo para HPS»: título, nemotécnico, código, clasificación, observaciones, contratistas, subcontratistas y lugares de ejecución. Es un formulario **no enlazado**: trabaja sobre el DTO global `m_ObjExpedienteDTOActivo` y guarda con `ExpedienteOperaciones.Registrar`. Las citas son `Form_FormExpedienteAltaParaHPS.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedienteAltaTipo` (`ComandoAlta_Click`) | Con el tipo `EXPHPS` crea un DTO nuevo, fija `TIpo` y los indicadores de tipo, cierra el formulario si estaba abierto, lo abre y se suscribe con `WithEvents` | `Form_FormExpedienteAltaTipo.cls:11, 88-107` |

`SetDatosDeExpedienteTipo` deja para `EXPHPS` los indicadores `EsAM = No`, `EsLote = No`, `EsBasado = No`, `EsExpediente = Sí` (`FUNCIONES UTILES.bas:2091-2095`). El título «ALTA DE EXPEDIENTE SÓLO PARA HPS» está fijado en el diseño (`.form.txt:612-613`).

Salidas: `CmdSalir_Click` cierra sin guardar (`:47-49`); `ComandoRegistrar_Click` cierra tras guardar (`:409`). Evento hacia quien lo abrió: `AltaExpediente` (`:6`, lanzado en `:406`). `Form_FormExpedienteAltaTipo` solo tiene manejador para el evento de `Form_FormExpediente` (`Form_FormExpedienteAltaTipo.cls:343`), no para el de este formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Titulo`, `Nemotecnico`, `CodExp`, `Observaciones` | textos | — | Se copian al DTO con `Nz` al guardar | `:419-423`, `.form.txt:718, 778, 1862, 912` |
| `IdGradoClasificacion` | combo 2 columnas (ID oculto) | — | Rellenado desde `m_ObjEntorno.GradosClasificaciones`; el DTO toma `.Column(0)` | `:84-93, 422`, `.form.txt:877-880` |
| `Contratistas` | combo 2 columnas | `AfterUpdate` | Añade la empresa elegida como contratista principal | `:431-433`, `.form.txt:1360-1365` |
| `ComandoAltaContratista` | botón | `OnClick` | `AgregarSuministradorMemoria` con rol `ROOT_CONTR`; repinta la lista y vacía el combo | `:195-210` |
| `ListaContratistas` | lista | — | Muestra los nodos con `ContratistaPrincipal = "Sí"` | `:131-152` |
| `ComandoEliminarContratista` | botón | `OnClick` | Quita del árbol el nodo seleccionado con rol `ROOT_CONTR` | `:212-226` |
| `SubContratistas` (`subContratistas` en el diseño) | combo 2 columnas | `AfterUpdate` | Añade la empresa como subcontrata directa | `:434-436`, `.form.txt:1931-1936` |
| `ComandoAltaSubContratista` | botón | `OnClick` | `AgregarSuministradorMemoria` con rol `ROOT_SUB` | `:232-247` |
| `ListaSubContratistas` | lista | — | Nodos con `SubContratista = "Sí"` y sin padre | `:154-175` |
| `ComandoEliminarSubContratista` | botón | `OnClick` | Quita el nodo con rol `ROOT_SUB` | `:249-262` |
| `LugaresEjecucion` | combo 2 columnas | `AfterUpdate` | Añade el lugar elegido | `:437-439`, `.form.txt:1071-1098` |
| `ComandoAltaLugarEjecucion` | botón | `OnClick` | Lee el lugar con `constructor.getLugarEjecucion` y lo añade a `ColLugaresEjecucion` si no estaba | `:340-359` |
| `ListaLugaresEjecucion` | lista | — | Muestra `ColLugaresEjecucion` | `:177-189` |
| `ComandoEliminarLugarEjecucion` | botón | `OnClick` | Quita el lugar seleccionado | `:361-372` |
| `ComandoRegistrar` (rótulo «Editar», aviso «Registrar») | botón | `OnClick` | Guarda el alta | `:378-415`, `.form.txt:2311-2323` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:47-49`, `.form.txt:2394-2396` |
| `ComandoAyuda` (rótulo «Salir»), `ComandoNuevaClasificacion`, `ComandoNuevoSuministrador`, `ComandoNuevoSuministrador1`, `ComandoNuevoLugar` | botones | sin `OnClick` en el diseño ni handler en el `.cls` | Sin comportamiento identificado | `.form.txt:639-640, 970, 1617, 1708, 2188` |

Los `RowSource` de diseño de los combos y listas (`.form.txt:879, 1073, 1362, 1493, 1933, 2064`) contienen datos de ejemplo; en ejecución se vacían y se rellenan desde las cachés de `Entorno` o desde el DTO (`:85, 97-98, 111-112, 137, 160, 182`).

## Flujos

### Apertura (`Form_Load`, `:20-45`)

1. `Ajustar Me` redimensiona la ventana (`FUNCIONES UTILES.bas:2703`) y fija `AllowEdits = True` sin consultar rol ni modo (`:26-27`).
2. Si el DTO no tiene `ColArbolSuministradores`, crea el diccionario (`:30-33`); `ExpedienteDTO` ya lo crea en `Class_Initialize` (`classes/ExpedienteDTO.cls:35-38`).
3. `EstablecerDatos`: con DTO `Nothing`, error «Contexto perdido (DTO Nothing)» (`:58-61`); rellena combos (`:63`) y las tres listas desde el DTO (`:67-69`).
4. Cualquier error muestra «Error al cargar Alta HPS: …» y cierra el formulario (`:41-44`).

### Composición en memoria

- **Alta de empresa** (`AgregarSuministradorMemoria`, `:268-305`): si el `IDSuministrador` ya está en el árbol con cualquier rol, avisa «Esta empresa ya está en la lista.» y no añade (`:274-280`). Si no, crea un `ExpedienteSuministrador` con clave temporal `TMP_<hhmmss>_<aleatorio>` (`:283`), `IdPadre = Null`, `Tag` = clave temporal (`:287-290`) y, según el rol, `ContratistaPrincipal`/`SubContratista` y la descripción «Contratista Principal» o «Participación Telefónica (HPS)» (`:292-301`).
- **Baja de empresa** (`EliminarSuministradorMemoria`, `:307-334`): borra el primer nodo con ese ID cuyo rol coincida.
- **Lugares**: diccionario con clave `IDLugarEjecucion`; los repetidos se ignoran sin aviso (`:354-357`).

### Guardado (`ComandoRegistrar_Click`, `:378-415`)

1. `RellenarDTO` copia cinco campos (`:388, 417-425`).
2. Fuerza `Ambito = "HPS"`, `HPSAplica = "Sí"`, `POSTAGEDO = "Sí"` y `TIpo = "Expediente Sólo para HPS"` (`:391-396`).
3. `ExpedienteOperaciones.Registrar` con `p_ExpedienteAlInicio = Nothing`, es decir, como alta (`:400-401`; `classes/ExpedienteOperaciones.cls:54`).
4. Éxito: lanza `AltaExpediente` y cierra (`:406-409`). Error: «Error al registrar: <m_Error>» y el formulario sigue abierto (`:412-414`).

## Reglas, validaciones y cálculos

- No hay validación propia en el formulario: la aplica `ExpedienteOperaciones.MotivoNoOK` (`classes/ExpedienteOperaciones.cls:61-64`). Entre otras, exige título (`:819-822`), `APLICAESTADO` «Sí» o «No» (`:839-842`), `CodExp` salvo lotes (`:894-897`), `AGEDYSAplica` informado (`:907-910`), unicidad de título, `CodExp` y nemotécnico (`:936-1001`) y las invariantes de responsables cuando hay `ColResponsables` (`:1037-1042, 1104`).
- Existe `MotivoNoOKSoloHPS`, con reglas propias de este tipo (al menos una empresa y un lugar de ejecución, `classes/ExpedienteOperaciones.cls:1178-1197`), pero ninguna línea de `src` la invoca.
- `Registrar` recalcula `TIpo = TipoCalculadoTexto` (`classes/ExpedienteOperaciones.cls:82`); con `EsExpediente = "Sí"`, `getTipoCalculado` devuelve `EXPIndividual` (`FUNCIONES UTILES.bas:5434-5436`).
- Una empresa no puede figurar a la vez como contratista y subcontratista (`:274-280`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combos de empresas | `TbSuministradores` vía caché `Entorno.Suministradores` | SQL dinámico | — | `classes/Entorno.cls:1016-1030`; `constructor.bas:3013-3014` |
| Nombre de cada empresa del árbol | `TbSuministradores` (una lectura por nodo) | SQL dinámico | — | `:146, 169`; `constructor.bas:5574-5591` |
| Combo y alta de lugares | `TbLugaresEjecucion` vía caché y `getLugarEjecucion` | SQL dinámico | — | `classes/Entorno.cls:1045-1059`; `constructor.bas:637-642` |
| Cabecera | `TbExpedientes` (`AddNew`) | DAO | `ws.BeginTrans` … `CommitTrans` | `classes/ExpedienteOperaciones.cls:47, 76-92, 520` |
| Lugares | `TbExpedientesLugaresEjecucion` (`AddNew`) | DAO | misma | `classes/ExpedienteOperaciones.cls:479-486, 1406-1410` |
| Árbol de empresas | `PersistirArbolDesdeMemoria` | — | misma | `classes/ExpedienteOperaciones.cls:489-495`; `classes/ExpedienteSuministradorServicio.cls:312-316` |
| Caché | `RegistrarExpEntidades` (ámbito `Cabecera`) | DAO | misma | `classes/ExpedienteOperaciones.cls:514` |

El resto del alta (NC, hijos vacíos, hook E2E, `PintarTareas`) se describe en [Form_FormExpediente](Form_FormExpediente.md#lecturas-y-escrituras).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la clave temporal del árbol es texto generado en VBA (`:283`) y debe mapearse a un identificador definitivo; el `IDExpediente` se calcula en VBA (`classes/ExpedienteOperaciones.cls:77`); cabecera y lugares van en la misma transacción DAO.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `PersistirArbolDesdeMemoria` es un cuerpo vacío con comentarios (`classes/ExpedienteSuministradorServicio.cls:312-316`): los contratistas y subcontratistas elegidos no se escribirían.
- El formulario nunca rellena `APLICAESTADO` ni `AGEDYSAplica` y `Expediente` no tiene `Class_Initialize` que los inicialice (`classes/Expediente.cls:46, 54`): `MotivoNoOK` rechazaría el alta con «APLICAESTADO se ha de rellenar» (`classes/ExpedienteOperaciones.cls:839-842`).
- `ComandoRegistrar` está deshabilitado en el diseño (`Enabled = NotDefault`, `.form.txt:2312`) y el `.cls` no lo habilita en ningún punto.
- El `TIpo` «Expediente Sólo para HPS» fijado en `:395` se sobrescribe con el texto de `EXPIndividual` (`classes/ExpedienteOperaciones.cls:82`; `FUNCIONES UTILES.bas:5434-5436`).
- `MotivoNoOKSoloHPS` (`classes/ExpedienteOperaciones.cls:1131`) no se usa: no se exige empresa ni lugar de ejecución.
- Nadie escucha `AltaExpediente` (`:406`): `Form_FormExpedienteAltaTipo` declara la variable `WithEvents` pero no su manejador (`Form_FormExpedienteAltaTipo.cls:11, 343`).
- Las variables `WithEvents` de formularios auxiliares nunca se asignan y los botones `ComandoNuevo*` no tienen `OnClick` (`:9-12, 442-454`; `.form.txt:970, 1617, 1708, 2188`): los manejadores de alta rápida son código muerto.
- Con un error no controlado en el guardado, el mensaje muestra `m_Error` vacío y no `Err.Description` (`:412-414`); en la carga ocurre lo contrario y se pierde el texto de `m_Error` (`:36, 43`).
- `ComandoAltaLugarEjecucion_Click` y `ComandoEliminarLugarEjecucion_Click` no tienen gestión de errores (`:340-372`).
- La clave temporal usa `Rnd` sin `Randomize` y la hora a segundos (`:283`); una colisión haría fallar `Dictionary.Add` (`:304`).
- Tras guardar no se limpian las globales del DTO, a diferencia de `Form_FormExpediente` (`Form_FormExpediente.cls:92-98`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: integrar el alta HPS en el alta general con un tipo explícito en lugar de un formulario aparte; validar con las reglas de `MotivoNoOKSoloHPS` y persistir el árbol de empresas en la misma transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. El alta HPS guarda `Ambito = HPS`, `HPSAplica = Sí` y `POSTAGEDO = Sí` (`:391-396`).
2. Añadir una empresa ya presente con cualquier rol muestra «Esta empresa ya está en la lista.» y no la duplica.
3. Las listas de contratistas y subcontratistas reflejan el árbol en memoria tras cada alta o baja.
4. Un lugar de ejecución repetido no se añade dos veces.
5. Cabecera, lugares y empresas se guardan en una sola transacción; si falla una parte, no queda nada escrito.
6. Un título o código ya existentes rechazan el alta con el mensaje de unicidad correspondiente.

## Pendientes

- Si el binario compilado conserva `Enabled = No` en `ComandoRegistrar` o lo habilita otra ruta no visible en `src`.
- Si el alta HPS ha funcionado alguna vez con este código: no hay prueba en `src/modules` que la ejerza (búsqueda de `AltaParaHPS` y `MotivoNoOKSoloHPS`).
- Implementación anterior de `PersistirArbolDesdeMemoria` y de `RecalcularFlagContratistaPrincipal`, citada solo en comentario (`classes/ExpedienteSuministradorServicio.cls:313-315`).
- Comportamiento previsto de `ComandoAyuda` y de los botones `ComandoNuevo*`.
