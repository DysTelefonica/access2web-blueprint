# FUNCIONES UTILES.bas — Copiar y extender entidades

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/modules/FUNCIONES UTILES.bas` (6854 líneas) · **PRD:** [README](../README.md)

Familia de funciones que copia las entidades de un expediente a otro y las propaga a su descendencia. **No figuraba en el PRD.** Las citas son `FUNCIONES UTILES.bas:línea`.

| Función | Copia | Evidencia |
|---|---|---|
| `CopiarSuministradores` | `TbExpedientesSuministradores` | `:3357-3408` |
| `CopiarPecales` | `TbExpedientesPECAL` | `:3410-3472` |
| `CopiarCPVs`, `CopiarLugares`, `CopiarComerciales` | CPV, lugares de ejecución y comerciales | `:3473-3598`, `:3599-3662` |
| `CopiarRACs`, `CopiarResponsables` | RAC y responsables | `:3663-3724`, `:3725-3792` |
| `CopiarAnualidades` | `TbExpedientesAnualidades` | `:3793-3874` |
| `ExtenderEntidadesDeExpediente(origen, destino)` | Llama, en este orden, a `CopiarSuministradores`, `CopiarPecales`, `CopiarCPVs`, `CopiarLugares`, `CopiarComerciales`, `CopiarRACs` y `CopiarResponsables`, y corta si alguna deja un error. **No copia anualidades** (`CopiarAnualidades` solo se invoca desde `TEST.bas:93`) | `:3875-3923` |
| `ExtenderEntidades(expediente)` | Aplica la anterior a cada hijo y a cada nieto, y registra el último cambio de cada uno (`RegistrarUltimoCambio`) | `:3924-3972` |

## Reglas

- Cada `Copiar*` exige origen y destino informados y ya registrados («El expediente de origen/Destino parece que todavía no ha sido registrado»), `:3369-3379, 3421-3431`.
- **Sustituyen, no fusionan:** borran antes las filas del destino y después insertan las del origen con un identificador nuevo (`CopiarPecales`: `DELETE` por `IDExpediente` e `IDPECAL`, luego `INSERT` con `IDPECALExpedienteCalculado`, `:3434-3460`).
- `ExtenderEntidades` falla con «No hay un expediente registrado» o «El expediente no tiene descendencia» (`:3934-3944`); solo llega a dos niveles de descendencia.
- Ni los códigos de compra ni las anualidades se extienden.

## Posibles defectos (hipótesis, sin ejecutar)

- **`CopiarSuministradores` abre como recordset una sentencia `DELETE` del origen** (`:3386-3389`): la segunda sentencia es `DELETE * FROM TbExpedientesSuministradores WHERE IDExpediente=<origen>` y se pasa a `OpenRecordset`, no a `Execute`. Parece un `SELECT` mal escrito. En DAO una consulta de acción no se abre como recordset, así que la función fallaría siempre; si se ejecutara, borraría los suministradores del origen. En ambos casos `ExtenderEntidades` no copiaría los suministradores.
- Los `DELETE` y `INSERT` se concatenan sin parámetros y sin transacción: un error a medias deja el destino parcialmente copiado.
- `ExtenderEntidades` no tiene llamadores en `src` (la única referencia externa a la familia es `CopiarAnualidades` en `TEST.bas:93`), ni en los formularios; se ejecuta, si se ejecuta, desde el binario.

## Comportamiento que la migración debe conservar

Pendiente de decisión de negocio: no se sabe si «extender entidades a la descendencia» es una funcionalidad viva. Si lo es, debe definirse en web como una operación única y transaccional, con la semántica de sustituir y con las anualidades y los códigos de compra decididos explícitamente.

## Pendientes

- Preguntar a negocio si se usa y desde qué opción de menú ([P-EXP-13](../issues-propuestas.md#p-exp-13--extender-entidades-a-la-descendencia)).
- Confirmar el defecto de `CopiarSuministradores` contra el binario.
