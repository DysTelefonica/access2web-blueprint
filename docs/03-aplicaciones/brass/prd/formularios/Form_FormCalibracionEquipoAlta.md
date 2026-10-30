# Form_FormCalibracionEquipoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionEquipoAlta.cls` (188 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un equipo calibrable: BUI (obligatorio, de una lista), equipo (obligatorio), modelo, número de serie y observaciones. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); solo el combo `ComboBui` tiene origen de filas (`TbBUI`, `Form_FormCalibracionEquipoAlta.form.txt:635`). Escribe una fila en `TbEquiposCalibrables` con DAO. Las citas son `Form_FormCalibracionEquipoAlta.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalibracionEquipoGestion` | cierra el formulario si estaba abierto y lo abre sin argumentos | `Form_FormCalibracionEquipoGestion.cls:131-142` |

No recibe `OpenArgs`. Al guardar refresca `FormCalibracionEquipoGestion` (`ActualizarLista`) si está abierto y se cierra (`:33-36`). `cmdSalir` cierra sin guardar (`:48-51`).

## Controles y eventos

Controles: 21 (30 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en cabecera, 1 `Section` y 21 controles de la sección Detalle).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboBui` | combo | — | BUI; obligatorio. Filas: `SELECT TbBUI.BUI … ORDER BY TbBUI.BUI` | CAP-5 | `:14`, `.form.txt:633-635` |
| `Equipo` | texto | — | Nombre del equipo; obligatorio | CAP-5 | `:11`, `.form.txt:463` |
| `Modelo` | texto | — | Opcional | CAP-5 | `:12`, `.form.txt:413` |
| `NS` | texto | — | Número de serie, opcional | CAP-5 | `:13`, `.form.txt:491` |
| `Observaciones` | texto | — | Opcional | CAP-5 | `:15`, `.form.txt:661` |
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida y registra el equipo | CAP-5 | `:4-47`, `.form.txt:356-358` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-5 | `:48-51`, `.form.txt:382-384` |
| `ComandoAyudaFormulario` («Comando46») | botón | sin evento en el `.form.txt` | Sin handler | sin asignar | `.form.txt:753-754` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «ALTA DE EQUIPO CALIBRABLE»; `lblTitulo1` pasa a `Caption` | CAP-5 | `:58`, `.form.txt:713-714, 733-734` |
| `Etiqueta538`, `Etiqueta222`, `Etiqueta225`, `Etiqueta234`, `Etiqueta238`, `Etiqueta1`, `Etiqueta2` | etiquetas | — | Rótulos (EQUIPO, MODELO, NÚM. SERIE, BUI, OBS.) y dos asteriscos | CAP-5 | `.form.txt:518, 545, 572, 599, 688, 792, 814` |
| `Cuadro69`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:288, 311, 334` |
| `Marco175` | grupo de opciones | — | Marco decorativo sin eventos | sin asignar | `.form.txt:439` |

## Flujos

### Apertura (`Form_Load`, `:52-64`)

Con `On Error Resume Next`: cursor de espera, `AjustarTamaño Me` (`FG.bas:261-288`), copia del título, repintado y fin de espera. No carga datos.

### Registro (`ComandoAlta_Click`, `:4-47`)

1. Lee los cinco controles con `Nz(…, "")` (`:11-15`) y llama a `AltaEquipoACalibrar(BUI, Equipo, Modelo, NS, Observaciones)` (`:21`; `:65-183`).
2. Resultado con `|` ⇒ error 1000 con el texto devuelto (`:22-26`).
3. Éxito: refresca la lista de gestión si está abierta y cierra (`:33-36`). Error: `MsgBox` crítico; el formulario queda abierto (`:41-46`).

## Reglas, validaciones y cálculos

- BUI vacío ⇒ «Es obligatorio indicar el BUI»; equipo vacío ⇒ «Es obligatorio indicar el Equipo» (`:88-95`).
- Duplicado: busca un equipo con el mismo `Equipo` y `BUI` y, según cuáles de `Modelo` y `NS` vengan informados, también por ellos (cuatro variantes del SQL, `:96-120`); si existe, «Para <BUI> el equipo <Equipo> ya ha sido registrado previamente» (`:122-126`).
- Identificador: `DameID("TbEquiposCalibrables","IDEquipoCalibrable")` = máximo + 1 (`:135-145`; `FG.bas:371-439`, cálculo en `FG.bas:428`); se comprueba que no exista ya (`:146-154`).
- Solo se escriben `Modelo`, `NS` y `Observaciones` si no están vacíos (`:159-167`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de BUI | `TbBUI` | `RowSource` guardado en el combo | — | `.form.txt:635` |
| Duplicado | `TbEquiposCalibrables` | SQL dinámico concatenado con textos entre comillas simples, DAO | no | `:96-121` |
| Siguiente ID | `TbEquiposCalibrables` | `DameID` (SELECT del campo y bucle en VBA) | no | `:135`; `FG.bas:371-439` |
| Comprobación de ID libre | `TbEquiposCalibrables` | SQL dinámico concatenado con el id | no | `:146-149` |
| Alta | `TbEquiposCalibrables` | DAO `AddNew`/`Update` (`IDEquipoCalibrable`, `Equipo`, `BUI` y los opcionales informados) | no | `:155-168` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID calculado en VBA con máximo + 1 y sin transacción; SQL concatenado con textos libres entre comillas simples; comparación de texto sujeta a la configuración de la base (`Option Compare Database`, `:2`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Equipo, modelo o NS con apóstrofo rompen el SQL concatenado (`:99-119`) y el alta falla con el error genérico.
- La clave de duplicado depende de qué campos vengan vacíos (`:96-120`): dos equipos con el mismo nombre y BUI pero uno con modelo y otro sin él podrían registrarse como distintos o como duplicados según la variante: a verificar.
- Si solo `Modelo` o solo `NS` vienen informados se usa una variante distinta de la consulta, y la que no se informa no participa (`:108-119`).
- `DameID` + `AddNew` sin transacción (`:135-168`): dos altas simultáneas podrían calcular el mismo ID.
- `ComboBui` acepta el valor de la lista; el código no comprueba que el BUI exista en `TbBUI` (`:88-91`); depende de la propiedad `LimitToList` del combo, no verificada.
- `Form_Load` con `On Error Resume Next` oculta errores de apertura (`:53`).
- `ComandoAyudaFormulario` carece de evento (`.form.txt:753-754`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: identificador generado por la base de datos y restricción única sobre la clave de equipo.

## Criterios de aceptación

Propuestos, no aprobados:

1. «Registrar» sin BUI o sin equipo muestra el mensaje de obligatoriedad y no escribe (`:88-95`).
2. Registrar un equipo con el mismo BUI y equipo (y mismos modelo/NS informados) muestra «ya ha sido registrado previamente» (`:122-126`).
3. Un alta correcta añade la fila con solo los campos informados y cierra el formulario (`:155-168, 36`).
4. Tras el alta con `FormCalibracionEquipoGestion` abierto, su lista se actualiza (`:33-35`).
5. `ComboBui` ofrece los BUI de `TbBUI` ordenados (`.form.txt:635`).
6. «Salir» cierra sin escribir (`:48-51`).

## Pendientes

- Si `ComboBui` limita a la lista (`LimitToList`) y claves/índices de `TbEquiposCalibrables` (ERD de origen).
- Contenido de `Form_FormCalibracionEquipoGestion.ActualizarLista` (`Form_FormCalibracionEquipoGestion.cls:33`).
