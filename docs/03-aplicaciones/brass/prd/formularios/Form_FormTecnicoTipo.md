# Form_FormTecnicoTipo

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoTipo.cls` (112 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un perfil de técnico (`TbTipoTecnico`) junto con su primera tarifa (`TbTipoTecnicoPrecios`: fecha inicial, precio hora laborable y precio hora extra). Es un formulario no enlazado que persiste con las clases `TIPOTECNICO` y `TipoTecnicoPrecio` (`Form_FormTecnicoTipo.cls:41-61`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `TipoTecnico.cls` = `classes/TipoTecnico.cls`; `TipoTecnicoPrecio.cls` = `classes/TipoTecnicoPrecio.cls`. Las citas son `Form_FormTecnicoTipo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTipoTecnicoPreciosGestion` | `ComandoAltaPerfil_Click`: cierra el formulario si estaba abierto, lo abre y se suscribe a su evento `Alta` | `Form_FormTipoTecnicoPreciosGestion.cls:11-20` |

Evento hacia quien lo abrió: `Alta(m_ObjTipoTecnicoPrecio)` (`:5`, lanzado en `:62`); el gestor lo recibe y refiltra la lista (`Form_FormTipoTecnicoPreciosGestion.cls:430-447`). Tras guardar se cierra (`:66`); `cmdSalir` cierra sin guardar (`:83-86`).

## Controles y eventos

Controles: 24 (24 Begin en el .form.txt: 1 formulario, 7 plantillas de control por defecto, 1 sección `Detalle` y 15 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `TIPO` | texto | — | Nombre del perfil; obligatorio | CAP-9 | `:25-28`, `.form.txt:402-427` |
| `FechaInicial` | texto | — | Fecha inicial de la tarifa; fecha corta | CAP-9 | `:29-32`, `.form.txt:494-521` |
| `PrecioHoraLab` | texto | — | Precio hora laborable; formato moneda | CAP-9 | `:33-36`, `.form.txt:604-633` |
| `PrecioHoraExtra` | texto | — | Precio hora extra; formato moneda | CAP-9 | `:37-40`, `.form.txt:635-669` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Valida, crea perfil y tarifa | CAP-9 | `:12-81`, `.form.txt:322-346` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-9 | `:83-86`, `.form.txt:348-373` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «PERFIL TÉCNICO»; `lblTitulo1` da el título de la ventana | CAP-9 | `:99`, `.form.txt:429-470` |
| `Etiqueta5`, `Etiqueta10`, `Etiqueta13`, `Etiqueta14` | etiquetas | — | Rótulos «Tipo Técnico», «Fecha Inicial», «€ Hora Lab», «€ Hora Extra» | CAP-9 | `.form.txt:375-400, 523-602` |
| `Etiqueta1` | etiqueta «*» | — | Marca de obligatorio | CAP-9 | `.form.txt:472-492` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-9 | `.form.txt:276-320` |

## Flujos

### Apertura (`Form_Load`, `:90-112`)

Ajusta el tamaño (`AjustarTamaño`, `FG.bas:261`) y copia el título a `Me.Caption` (`:97-99`). Ante error: `MsgBox` con el texto y cierre del formulario (`:105-111`).

### Guardado (`ComandoRegistrar_Click`, `:12-81`)

1. Validaciones de formulario (ver abajo) con `m_Error` y `Err.Raise 1000` (`:25-40`).
2. Crea `TIPOTECNICO` con `ParaAlta = Sí` y el texto de `TIPO`, y lo registra (`:41-44`); marca `blnTipoDadoDeAlta` (`:48`).
3. Crea `TipoTecnicoPrecio` con el perfil, fecha y precios, y lo registra (`:49-61`).
4. Lanza `Alta` y cierra (`:62-69`).
5. Error: si el perfil ya se había creado, lo borra con `Borrar` (`:72-74`); muestra `m_Error` en un `MsgBox` crítico y el formulario queda abierto (`:76-80`).

## Reglas, validaciones y cálculos

- Formulario: `TIPO` vacío ⇒ «Falta tipo técnico»; fecha inicial no válida; precios no numéricos (`:25-40`).
- Perfil: «Obligatorio el Tipo» (`TipoTecnico.cls:229-232`); en alta, si el `TIPO` ya existe en `TbTipoTecnico` ⇒ «El TIPO ya ha sido usado» (`TipoTecnico.cls:254-262`).
- Tarifa: en alta no admite fecha final (`TipoTecnicoPrecio.cls:319-323`); precios numéricos (`TipoTecnicoPrecio.cls:311-318`).
- Cierre de la tarifa vigente: si ya existe una tarifa sin fecha final para el perfil, se le pone la fecha de hoy (`TipoTecnicoPrecio.cls:431-440`); en caso de error se restaura a `Null` (`TipoTecnicoPrecio.cls:511-515`).
- Identificador de la tarifa: `DameID("TbTipoTecnicoPrecios", "IDTipoTecnicoFecha")` (`TipoTecnicoPrecio.cls:108`), máximo + 1 (`FG.bas:408-428`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Alta del perfil | `TbTipoTecnico` | DAO `AddNew`/`Update` | no | `TipoTecnico.cls:361, 377-380` |
| Búsqueda de perfil existente | `TbTipoTecnico` | `ValorExistente` (comprobación de valor) | — | `TipoTecnico.cls:255` |
| Tarifa vigente | `TbTipoTecnicoPrecios` | SQL dinámico con el perfil concatenado | — | `Constructor.bas:2211-2215` |
| Cierre de tarifa vigente | `TbTipoTecnicoPrecios` | SQL dinámico `UPDATE` con fecha literal `mm/dd/yyyy` | no | `TipoTecnicoPrecio.cls:436-438` |
| Alta de la tarifa | `TbTipoTecnicoPrecios` | DAO `AddNew`/`Update` | no | `TipoTecnicoPrecio.cls:470-485` |
| Compensación por error | `TbTipoTecnico` | SQL dinámico `DELETE` | no | `TipoTecnico.cls:452-454` |

## Coexistencia

Adaptar a PostgreSQL vinculado: fecha literal `#mm/dd/yyyy#` (`TipoTecnicoPrecio.cls:436`) y `DELETE *` (`TipoTecnico.cls:452`) en dialecto Access; máximo + 1 en VBA para el ID de tarifa; formato de moneda de los controles (`.form.txt:617-618, 648-649`); la compensación manual sustituiría a una transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Sin transacción: el perfil se crea antes que la tarifa y se compensa con `Borrar` (`:72-74`); si `Borrar` falla, `m_Error` se sobrescribe y puede quedar un perfil huérfano.
- Si el perfil se crea y la tarifa falla tras cerrar una tarifa previa, depende de la restauración de `TipoTecnicoPrecio.cls:511-515`.
- `m_Error` se mezcla con `p_Error` en `TipoTecnicoPrecio.cls:447` (lectura de `Me.Error`).
- `Form_Load` anuncia «Al Form_Open» en su mensaje (`:107`).
- SQL concatenado con el perfil entre comillas (`Constructor.bas:2213`, `TipoTecnico.cls:453`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: alta de perfil y tarifa en una única transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Registrar sin tipo, sin fecha inicial o con un precio no numérico muestra el mensaje correspondiente y no escribe (`:25-40`).
2. Registrar un tipo ya existente muestra «El TIPO ya ha sido usado» (`TipoTecnico.cls:259-262`).
3. Un alta válida crea perfil y tarifa, lanza `Alta` y cierra (`:41-69`).
4. Si falla la tarifa, el perfil recién creado se elimina (`:72-74`).
5. Al abrir desde la gestión, el listado se refiltra tras el alta (`Form_FormTipoTecnicoPreciosGestion.cls:430-447`).

## Pendientes

- Contenido de `ValorExistente` y alcance de `Constructor.getTipoTecnicoActivo` más allá de `Constructor.bas:2211-2234`.
- Longitudes y claves de `TbTipoTecnico` y `TbTipoTecnicoPrecios` (ERD de origen).
- `Borrar` puede rechazar el borrado si el tipo tiene facturas o técnicos (`TipoTecnico.cls:416-430`), no probado.
