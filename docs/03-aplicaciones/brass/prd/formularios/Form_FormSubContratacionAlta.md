# Form_FormSubContratacionAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubContratacionAlta.cls` (161 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de una subcontratación en `TbSubcontrataciones` (empresa, fecha de alta, descripción, importe y fecha de fin opcional). Es un formulario no enlazado: los datos se leen de los controles y se persisten con la función pública `Alta` del propio formulario (`Form_FormSubContratacionAlta.cls:66-156`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormSubContratacionAlta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSubContratacionGestion` | Cierra el formulario si estaba abierto y lo abre sin argumentos | `Form_FormSubContratacionGestion.cls:105-108` |

Salida: tras guardar, si `FormSubContratacionGestion` está abierto llama a su `ActualizarLista` (`:32-34`) y se cierra con `acSaveNo` (`:35`). `cmdSalir` cierra sin guardar (`:48-51`). No lanza eventos propios.

## Controles y eventos

Controles: 29 (29 Begin en el .form.txt: 1 formulario, 7 plantillas de control por defecto, 1 sección `Detalle` y 20 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboEmpresa` | combo | — | Empresa subcontratada; lista las empresas ya registradas (`SELECT DISTINCT`) | CAP-8 | `.form.txt:539-574` |
| `FechaAlta` | texto | — | Fecha de alta; formato fecha corta | CAP-8 | `.form.txt:510-537` |
| `Descripcion` | texto | — | Descripción opcional | CAP-8 | `.form.txt:630-656` |
| `Importe` | texto | — | Importe; formato moneda con dos decimales | CAP-8 | `.form.txt:685-713` |
| `FechaFin` | texto | — | Fecha de fin opcional; formato fecha corta | CAP-8 | `.form.txt:742-769` |
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida y da de alta | CAP-8 | `:4-46`, `.form.txt:359-396` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-8 | `:48-51`, `.form.txt:398-436` |
| `Numero_Actividad` | texto bloqueado, deshabilitado, ancho 6 | — | Enlazado a `IDActividad`; sin uso en el código | sin asignar | `.form.txt:279-311` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Alta de Subcontratación»; `lblTitulo1` da el título de la ventana | CAP-8 | `:59`, `.form.txt:438-481` |
| `Etiqueta142`, `Etiqueta198`, `Etiqueta201`, `Etiqueta219`, `Etiqueta222` | etiquetas | — | Rótulos «Fecha Alta», «Empresa», «Descripción», «Importe», «Fecha Fin» | CAP-8 | `.form.txt:483-770` |
| `Etiqueta4`, `Etiqueta1`, `Etiqueta2` | etiquetas «*» | — | Marcas de obligatorio (3: empresa, fecha, importe) | CAP-8 | `.form.txt:771-840` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-8 | `.form.txt:313-357` |

## Flujos

### Apertura (`Form_Load`, `:52-65`)

`On Error Resume Next`; ajusta el tamaño con `AjustarTamaño` (`FG.bas:261`) y copia el título a `Me.Caption` (`:59`). Los errores se ignoran.

### Alta (`ComandoAlta_Click`, `:4-46`)

1. Lee los cinco controles con `Nz(…, "")` como texto (`:10-15`) y llama a `Alta` (`:21`).
2. `Alta` devuelve el ID, o `#ERR|texto` (`:155`); si el resultado contiene «|» se muestra el texto con `Err.Raise 1000` (`:22-26`).
3. Actualiza la lista de gestión si está abierta, cierra y quita el reloj de arena (`:32-38`).
4. Error: `MsgBox` crítico y el formulario queda abierto (`:40-46`).

## Reglas, validaciones y cálculos

- Empresa vacía ⇒ «Se ha de introducir el nombre de la empresa subcontratada» (`:87-90`).
- Fecha de alta no válida (`IsDate`) ⇒ «Se ha de introducir una fecha inicial válida» (`:91-94`).
- Importe no numérico ⇒ «Se ha de introducir un importe válido» (`:95-98`).
- Fecha de fin informada y anterior a la de alta ⇒ error (`:99-104`).
- Identificador: `DameID("TbSubcontrataciones", "IDSubcontratacion")` (`:110`) recorre la tabla y devuelve el máximo + 1 (`FG.bas:408-428`); se comprueba que no exista ya (`:121-129`).
- Descripción, fecha de fin e importe solo se escriben si traen valor (`:134-142`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de empresas | `TbSubcontrataciones` | `RowSource` guardada en el control (SQL con `DISTINCT`) | — | `.form.txt:559-561` |
| Siguiente ID | `TbSubcontrataciones` | SQL dinámico, recorrido de todos los registros | no | `FG.bas:371-428` |
| Comprobación de existencia | `TbSubcontrataciones` | SQL dinámico con el ID concatenado | no | `:121-124` |
| Alta | `TbSubcontrataciones` | DAO `AddNew`/`Update` | no | `:130-143` |

## Coexistencia

Adaptar a PostgreSQL vinculado: ID calculado en VBA con máximo + 1 (sustituible por secuencia); `Format="Short Date"` y formato de moneda del control (`.form.txt:523, 698`); asignación de cadenas a campos de fecha e importe (`:132, 141`), que dependerá de la configuración regional.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Máximo + 1 sin transacción ni bloqueo (`:110`, `FG.bas:408-428`): dos altas simultáneas podrían obtener el mismo ID.
- El mensaje de error de `Alta` empieza por «El método Eliminar ha devuelto un error» (`:24`), texto heredado de otro formulario.
- Si `FechaFin` trae un texto que no es fecha, `CDate` falla (`:100`) y el usuario ve el mensaje genérico de `Alta` (`:148-150`) en lugar de uno de validación.
- No se validan longitudes ni importes negativos; solo `IsNumeric` (`:95`).
- `Form_Load` ignora todos los errores (`:53`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos; validación de importe no negativo; lista cerrada o libre de empresas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Registrar sin empresa muestra «Se ha de introducir el nombre de la empresa subcontratada» y no escribe (`:87-90`).
2. Registrar con fecha de alta no válida o importe no numérico muestra el mensaje correspondiente (`:91-98`).
3. Con fecha de fin anterior a la de alta se rechaza (`:99-104`).
4. Un alta válida crea el registro con ID máximo + 1, cierra el formulario y refresca la lista de gestión si estaba abierta (`:116-145`, `:32-35`).
5. Descripción, fecha de fin e importe vacíos no se escriben (`:134-142`).
6. `Salir` cierra sin guardar (`:48-51`).

## Pendientes

- Si `ComboEmpresa` admite texto fuera de la lista (`LimitToList` no aparece en `.form.txt:539-574`).
- Si `Numero_Actividad` tiene algún uso real (`.form.txt:292`).
- Índices, tipos y longitudes de `TbSubcontrataciones` (ERD de origen).
