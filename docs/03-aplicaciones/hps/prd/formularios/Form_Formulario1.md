# Form_Formulario1

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_Formulario1.cls` (27 líneas) y `.form.txt` (1 273 líneas) · **PRD:** [README](../README.md)

Objeto de pruebas con el nombre por defecto de Access, sin uso en `src/`. Su `.cls` declara `Form_Load` con el cuerpo casi todo comentado: las llamadas a `Debug.Print` de colores de fondo están en comentarios y solo queda el `End Sub` por defecto. En el `.form.txt` queda un esquema decorativo de maestro/esclavo (rectángulos y etiquetas) con tres botones (`ComandoEscritura`, `ComandoSoloLectura`, `ComandoEliminarUsuario`) y una etiqueta `lblTitulo` con el rótulo «CONSULTAS». No es subformulario de nadie ni es anfitrión de subformularios: `DefaultView =0` (`formulario simple`) y no tiene propiedades de `Subform`/origenes. Las citas son `Form_Formulario1.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| — | Ninguna referencia en `src/`: ni `OpenForm`, ni `SourceObject`, ni nombre en módulos o clases | búsqueda de `Formulario1` en `src/**/*.cls` y `*.bas` sin resultados |

No abre nada.

## Controles y eventos

`DefaultView =0` (formulario simple) (`.form.txt:8`). `OnLoad ="[Event Procedure]"` (`.form.txt:364`).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Rectángulos del esquema maestro/esclavo | `CuadroDetalleMaster`, `CuadroCampoMaster`, `CuadroEtiquetaCampoMaster`, `CuadroListaMaster`, `CuadroDetalleSlave`, `CuadroCampoSlave`, `CuadroEtiquetaCampoSlave`, `CuadroListaSlave` | `.form.txt:831, 853, 874, 895, 954, 975, 996, 1017` |
| Etiquetas del esquema | `EtiquetaLinkMaster`, `EtiquetaCampoMaster`, `EtiquetaLinkSlave`, `EtiquetaCampoSlave` | `.form.txt:916, 935, 1039, 1058` |
| Etiqueta de título | `lblTitulo` («CONSULTAS») | `.form.txt:1077` |
| Botones | `ComandoEscritura` («Pasar a Solo Lectura»), `ComandoSoloLectura` («Pasar a Solo Lectura»), `ComandoEliminarUsuario` («Eliminar», `ControlTipText` «Borrar todos los datos del Usuario») | `.form.txt:1105, 1148, 1192` |

El `.cls` no declara handler para ninguno de los tres botones: no hay `Sub` con sufijo `_Click`.

## Flujos

- **Apertura (`Form_Load`, `:6-26`):** la única sentencia activa es la salida del propio `Sub` (`:26`); las veinte líneas anteriores son `Debug.Print` comentados sobre los `BackColor`/`ForeColor`/`BorderColor` de los rectángulos y etiquetas del esquema maestro/esclavo (`:7-25`). No se ejecuta nada.
- **Clics de los botones:** sin handler en el `.cls` (`:5-27`); el código no intercepta nada.

## Reglas, validaciones y cálculos

Ninguna.

## Lecturas y escrituras

Ninguna. `RecordSource` no informado (`:1-27`) y el `.form.txt` no fija `RecordSource` en el formulario (búsqueda sin resultados).

## Coexistencia

Sin impacto (sin acceso a datos).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Objeto de pruebas huérfano: ningún código de `src/` lo abre. El rótulo «CONSULTAS» (`lblTitulo`, `Caption` en `.form.txt:1088`) y el `ControlTipText` «Borrar todos los datos del Usuario» (`ComandoEliminarUsuario`, `.form.txt:1203`) sugieren una copia de otra utilidad.
- `Debug.Print` como única lógica prevista (`:7-25`), todo comentado.
- Los dos botones comparten el mismo `Caption` «Pasar a Solo Lectura» (`ComandoEscritura` y `ComandoSoloLectura`, `.form.txt:1115, 1159`): confundirían al lector; falta el opuesto de `ComandoEscritura`.
- `BotonVerde` declarado a nivel de formulario (`:4`) sin uso.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no migrarlo.

## Criterios de aceptación

Ninguno: no aporta capacidad de negocio.

## Pendientes

- Confirmar que ningún menú en datos ni macro del binario (`.accdb`) lo abre; el nombre aparece solo en su propio `.cls`/`.form.txt`.
- Decidir en el inventario si se clasifica como **no aplicable**.
