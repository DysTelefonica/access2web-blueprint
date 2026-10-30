# Form_FormTecnicoFiestas

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoFiestas.cls` (79 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un día festivo del calendario de técnicos (`TbTecnicosFiestas`: fecha y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): crea un objeto `Fiesta` (`classes/Fiesta.cls`) y delega la validación y la escritura en `Fiesta.Registrar`. Las citas sin fichero son `Form_FormTecnicoFiestas.cls:línea`; `Fiesta.cls:línea` es `classes/Fiesta.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoFiestasGestion` (`ComandoAlta_Click`) | Abre el formulario y lo asigna a `m_Form` (`WithEvents`) para escuchar `Alta` | `Form_FormTecnicoFiestasGestion.cls:7, 92-96` |
| `Form_FormTecnicoGestion` | Abre `FormTecnicoFiestasGestion` (el listado), no este formulario | `Form_FormTecnicoGestion.cls:120` |

Evento hacia el llamador: `Alta` (`:7`, lanzado en `:28` tras cada alta correcta); el listado lo atiende con `m_Form_Alta` y recarga la lista con el año seleccionado (`Form_FormTecnicoFiestasGestion.cls:298-313`). `cmdSalir` cierra sin guardar (`:51-54`).

## Controles y eventos

Controles: 10 (10 `Begin` de controles dentro de la sección Detalle del `.form.txt`, líneas 255-486; excluye `Begin Form`, `Begin Section` y las plantillas por defecto).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `CuadroSuperior`, `Cuadro2` | rectángulos (2) | — | Marcos decorativos | CAP-9 Técnicos, ausencias, fiestas y precios | `.form.txt:255, 277` |
| `lblTitulo1`, `lblTitulo2` | etiquetas (2) | — | Rótulo «ALTA DE FIESTA»; `lblTitulo1` da título a la ventana | CAP-9 Técnicos, ausencias, fiestas y precios | `:64`, `.form.txt:464-496` |
| `Etiqueta3`, `Etiqueta5` | etiquetas (2) | — | Rótulos «Fecha» y «Fiesta» | CAP-9 Técnicos, ausencias, fiestas y precios | `.form.txt:382, 409` |
| `FechaFiesta` | cuadro de texto no enlazado, formato «Short Date» | — | Fecha del festivo; obligatoria | CAP-9 Técnicos, ausencias, fiestas y precios | `:22`, `.form.txt:354-366` |
| `Observaciones` | cuadro de texto no enlazado | — | Descripción del festivo; obligatoria (rótulo «Fiesta») | CAP-9 Técnicos, ausencias, fiestas y precios | `:23`, `.form.txt:436-448` |
| `ComandoRegistrar` (rótulo «Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Valida y guarda; ofrece otra alta | CAP-9 Técnicos, ausencias, fiestas y precios | `:9-50`, `.form.txt:301-311` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-9 Técnicos, ausencias, fiestas y precios | `:51-54`, `.form.txt:327-337` |

## Flujos

### Apertura (`Form_Load`, `:57-79`)

Cursor de espera, `AjustarTamaño Me`, título de ventana desde `lblTitulo1` y repintado (`:61-69`). Error: `MsgBox` crítico y cierre del formulario (`:72-78`).

### Alta (`ComandoRegistrar_Click`, `:9-50`)

1. Crea `Fiesta`, copia `FechaFiesta` y `Observaciones` con `Nz(..., "")` (`:19-23`) y llama a `Registrar m_Error` (`:24`).
2. Si `m_Error` no queda vacío ⇒ error controlado (1000) (`:25-27`).
3. Éxito: lanza `Alta` (`:28`) y pregunta «¿Desea introducir alguna fiesta más?» (`:35`). Con «No» cierra el formulario (`:36-39`); con «Sí» vacía los dos campos y devuelve el foco a `FechaFiesta` (`:40-42`).
4. Error: restablece el cursor y `MsgBox` crítico con `m_Error`; el formulario queda abierto con los datos tecleados (`:44-49`).

## Reglas, validaciones y cálculos

- **Fecha obligatoria y válida:** `IsDate(FechaFiesta)`; si no, «Se ha de indicar la fecha de la fiesta» (`Fiesta.cls:30-33`).
- **Descripción obligatoria:** `Observaciones` vacía ⇒ «Se ha de indicar una fiesta» (`Fiesta.cls:34-37`).
- **Unicidad por fecha:** cuenta filas de `TbTecnicosFiestas` con esa fecha; si hay alguna, «La fecha ya ha sido registrada previamente» (`Fiesta.cls:38-50`).
- **Orden:** primero fecha, luego descripción, luego duplicado (`Fiesta.cls:30-50`); `Registrar` solo escribe si no hay motivo (`Fiesta.cls:82-89`).
- **Texto de la fecha:** la fecha viaja como `String` en `Fiesta.FechaFiesta` (`Fiesta.cls:14`) y se asigna tal cual al campo (`Fiesta.cls:94`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Comprobar fecha repetida | `TbTecnicosFiestas` | SQL dinámico con literal de fecha `#mm/dd/yyyy#` (`Format`), vía `getdb()` | no | `Fiesta.cls:38-41` |
| Alta (`FechaFiesta`, `Observaciones`) | `TbTecnicosFiestas` | DAO `OpenRecordset` sobre la tabla, `AddNew`/`Update`, vía `CurrentDb()` | no | `Fiesta.cls:90-97` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literal de fecha Jet `#mm/dd/yyyy#` y `Format` en el SQL de comprobación (`Fiesta.cls:40`); apertura de tabla completa con `AddNew` (patrón DAO sin SQL); comprobación y alta separadas, sin transacción; dos accesos distintos a la base (`getdb()` en `Fiesta.cls:41`, `CurrentDb()` en `Fiesta.cls:91`), que habría que unificar sobre la misma conexión vinculada.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Carrera entre contar y escribir (`Fiesta.cls:38-97`): dos altas simultáneas de la misma fecha podrían duplicarla si no hay índice único.
- `MotivoNoOK` cierra `rcdDatos` tras el bloque `If` aunque sea `Nothing` (`Fiesta.cls:42-53`); si `getdb()` devolviera `Nothing` fallaría al cerrar.
- `FechaFiesta` se guarda como texto asignado a un campo de fecha (`Fiesta.cls:14, 94`), sensible a la configuración regional.
- Una fecha con hora o un texto fecha aceptado por `IsDate` pero no normalizado podría no coincidir en la comprobación de duplicados (`Fiesta.cls:30, 40`).
- Tras un alta correcta lanza `Alta` aunque el usuario vaya a introducir otra; el listado se recarga en cada una (`:28`; `Form_FormTecnicoFiestasGestion.cls:298-304`).
- `pregunta` es una global compartida (`Variables Globales.bas:204`) usada sin declarar en el formulario (`:35, 49`).
- Si el formulario se abre sin pasar por la gestión, `Alta` no tiene oyente y el listado no se refresca (`Form_FormTecnicoFiestasGestion.cls:7, 92-96`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única sobre la fecha; tipo fecha explícito; una sola operación de alta con confirmación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Registrar con la fecha vacía o inválida muestra «Se ha de indicar la fecha de la fiesta» y no escribe (`Fiesta.cls:30-33`).
2. Registrar con la descripción vacía muestra «Se ha de indicar una fiesta» (`Fiesta.cls:34-37`).
3. Registrar una fecha ya existente muestra «La fecha ya ha sido registrada previamente» (`Fiesta.cls:44-48`).
4. Un alta correcta añade una fila a `TbTecnicosFiestas` y el listado de gestión se recarga con el año elegido (`Fiesta.cls:93-96`; `Form_FormTecnicoFiestasGestion.cls:298-304`).
5. Responder «Sí» a «¿Desea introducir alguna fiesta más?» deja los campos vacíos con foco en la fecha; «No» cierra (`:35-42`).
6. Ante un error el formulario permanece abierto con los datos introducidos (`:44-49`).
7. `Salir` cierra sin escribir (`:51-54`).

## Pendientes

- Contenido de `AjustarTamaño` (`:63`; definición en `FG.bas:261`).
- Si `TbTecnicosFiestas` tiene índice único sobre `FechaFiesta` y tipos de columna (ERD de origen).
- Definición de `getdb` y su base de destino (`Variables Globales.bas:548`).
