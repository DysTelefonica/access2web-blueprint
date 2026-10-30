# Form_FormCalibracionAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormCalibracionAlta.cls` (228 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de una calibración (fecha de calibrado y fecha de fin) para un equipo calibrable concreto. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el identificador del equipo en `OpenArgs`, muestra sus datos de solo lectura y escribe una fila en `TbEquiposCalibrablesFechas` con DAO. Las citas son `Form_FormCalibracionAlta.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCalibracionGestion` | `ComandoAltaCalibracion_Click`: exige id numérico, cierra el formulario si estaba abierto y lo abre con el id en `OpenArgs` | `Form_FormCalibracionGestion.cls:125-143` |

`OpenArgs` es el `IDEquipoCalibrable` (`:71`). No devuelve datos; al guardar refresca `FormCalibracionEquipoGestion` (llamando a `ActualizarLista`) y la lista `ListaFiltrados` de `FormCalibracionGestion` si están abiertos, y se cierra (`:34-46`). `cmdSalir` cierra sin guardar (`:58-61`). Variable de módulo: `strIDEquipoCalibracion` (`:4`).

## Controles y eventos

Controles: 25 (34 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en cabecera, 1 `Section` y 25 controles de la sección Detalle).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaCalibracion` | texto, formato fecha corta | — | Fecha de calibrado; obligatoria al guardar | CAP-5 | `:15`, `.form.txt:409-410` |
| `FechaFinCalibracion` | texto, formato fecha corta | — | Fecha de fin de calibrado; obligatoria al guardar | CAP-5 | `:16`, `.form.txt:747-748` |
| `Equipo`, `Modelo`, `NS`, `BUI` | textos con `Enabled` y `Locked` distintos del valor por defecto | — | Datos del equipo, cargados en `Form_Load` solo si no están vacíos | CAP-5 | `:92-103`, `.form.txt:585-586, 599, 616-617, 630, 647-648, 661, 705-706, 719` |
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida y registra la calibración | CAP-5 | `:5-57`, `.form.txt:330-332` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-5 | `:58-61`, `.form.txt:356-358` |
| `ComandoAyudaFormulario` («Comando46») | botón | sin evento en el `.form.txt` | Sin handler | sin asignar | `.form.txt:867-868` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «ALTA DE CALIBRACIÓN»; `lblTitulo1` pasa a `Caption` del formulario | CAP-5 | `:70`, `.form.txt:827-828, 847-848` |
| `Etiqueta538`, `Etiqueta190`, `Etiqueta191`, `Etiqueta192`, `Etiqueta193`, `Etiqueta197`, `Etiqueta243`, `Etiqueta245`, `Etiqueta99`, `Etiqueta0` | etiquetas | — | Rótulos de campos y dos asteriscos de obligatoriedad (`Etiqueta99`, `Etiqueta0`) | CAP-5 | `.form.txt:437, 488, 514, 541, 568, 688, 775, 803, 906, 928` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:285, 308` |
| `Marco175`, `Marco189` | grupos de opciones | — | Marcos decorativos sin eventos | sin asignar | `.form.txt:385, 463` |

## Flujos

### Apertura (`Form_Load`, `:62-121`)

1. Cursor de espera y `AjustarTamaño Me` (`:67-69`; `FG.bas:261-288`).
2. `OpenArgs` no numérico ⇒ «No se ha abierto con los parámetros adecuados» (`:71-75`).
3. Lee `Equipo`, `Modelo`, `NS` y `BUI` de `TbEquiposCalibrables` por `IDEquipoCalibrable`; sin fila, mismo mensaje (`:76-89`) y copia a los controles los valores no vacíos (`:92-103`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:110-120`).

### Registro (`ComandoAlta_Click`, `:5-57`)

1. Id de equipo no numérico ⇒ «No se puede determinar el equipo» (`:11-14`).
2. Lee las dos fechas con `Nz(…, "")` y llama a `AltaCalibracion` (`:15-22`; `:122-219`).
3. Resultado con `|` ⇒ error 1000 con el texto devuelto (`:23-27`).
4. Éxito: refresca las dos listas si están abiertas, cierra el formulario (`:34-46`).
5. Error: `MsgBox` crítico; el formulario queda abierto (`:51-56`).

## Reglas, validaciones y cálculos

- Las dos fechas deben ser fechas válidas (`IsDate`): «Es obligatoria la fecha de calibrado» / «… de fin de calibrado» (`:146-153`).
- Fin ≥ inicio, si no «La fecha de fin de calibrado ha de ser igual o mayor que la de calibrado» (`:154-157`).
- Si `TbEquiposCalibrables.FechaBajaParaCalibracion` es una fecha, el equipo está fuera de calibración y no se admite el alta (`:163-173`).
- Identificador: `DameID("TbEquiposCalibrablesFechas","IDCalibracion")` = máximo + 1 (`:179-189`; `FG.bas:371-439`, cálculo en `FG.bas:428`). Se comprueba que no exista ya esa fila (`:190-198`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos del equipo | `TbEquiposCalibrables` | SQL dinámico concatenado con el id | — | `:76-79` |
| Fecha de baja | `TbEquiposCalibrables` | `Dame` (SQL dinámico concatenado) | no | `:163`; `FG.bas:290-330` |
| Siguiente ID | `TbEquiposCalibrablesFechas` | `DameID` (SELECT del campo y bucle en VBA) | no | `:179`; `FG.bas:371-439` |
| Comprobación de ID libre | `TbEquiposCalibrablesFechas` | SQL dinámico concatenado | no | `:190-193` |
| Alta | `TbEquiposCalibrablesFechas` | DAO `AddNew`/`Update` (`IDCalibracion`, `IDEquipoCalibrable`, `FechaCalibrado`, `FechaFinCalibrado`) | no | `:199-204` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID calculado en VBA con máximo + 1 y sin transacción (autonumérico en destino); las fechas viajan como texto (`String`) hasta el campo (`:202-203`), lo que depende de la configuración regional; SQL concatenado con el id numérico.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `DameID` + `AddNew` sin transacción (`:179-204`): dos altas simultáneas podrían calcular el mismo ID.
- Las fechas se asignan como cadena (`:202-203`) y se comparan con `CDate` (`:154`): el resultado puede depender del formato regional.
- No se comprueba solape con otras calibraciones del mismo equipo (`:122-219`): no hay lectura previa de las existentes.
- El mensaje de error del alta menciona «URLCompletaArchivoAyuda» (`:25`), nombre que no corresponde a este método.
- `ComandoAyudaFormulario` carece de evento (`.form.txt:867-868`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: identificador generado por la base de datos y validación de solape de periodos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin id de equipo numérico muestra «No se ha abierto con los parámetros adecuados» y cierra (`:71-75, 119-120`).
2. Los campos de equipo se muestran sin poder editarse (`.form.txt:585-586`).
3. Guardar sin fecha de calibrado o sin fecha de fin muestra el mensaje correspondiente y no escribe (`:146-153`).
4. Guardar con fin anterior al inicio muestra el mensaje de orden y no escribe (`:154-157`).
5. Guardar para un equipo con fecha de baja de calibración muestra «fuera de la calibración» y no escribe (`:170-173`).
6. Un alta correcta cierra el formulario y refresca las listas abiertas (`:34-46`).

## Pendientes

- Semántica de `FechaBajaParaCalibracion` y tipos de las columnas de fecha (ERD de origen).
- Contenido de `Form_FormCalibracionEquipoGestion.ActualizarLista` (`Form_FormCalibracionEquipoGestion.cls:33`).
