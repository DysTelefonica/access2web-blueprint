# Perfil sintáctico VBA — gestion_riesgos (input para codegraph-vba round-3/round-5)

**Autor:** consumer `gestion_riesgos` (issue #69), en sesión 2026-07-13.
**Repositorio origen:** `DysTelefonica/GESTION_RIESGOS` rama `staging` SHA `72325a1`.
**Archivos analizados:** 128 `.cls` + 141 `.bas` (269 archivos totales).
**Total líneas de código:** ~95k (mixto .cls y .bas).
**Generado por:** `tools/scratch-vba-syntax-profile.mjs` (script autocontenido, regex sobre contenido completo, comentarios VBA excluidos).

## Por qué este documento

El maintainer de `@aroman22/codegraph-vba` archivó rounds #108 (clasificación de `unresolved_refs.reference_kind`) y #110 (post-extraction resolver repointing). Para escribir los tests RED correctamente necesita **datos reales del codebase VBA del consumer**, no fixtures inventados. Este perfil es el input.

## Conteo de patrones sintácticos

| Patrón | .cls | .bas | Total | Significado para round-3 |
|---|---|---|---|---|
| `paren-call-bare` (Foo(...)) | 7,315 | 21,959 | **29,274** | Llamada de función/proc — categoría `call` |
| `paren-call-qualified` (.Foo(...)) | 1,756 | 2,830 | **4,586** | Llamada sobre instancia — categoría `qualified-call` |
| `property-get` (.Foo lectura) | 25,759 | 15,593 | **41,352** | Lectura de propiedad — categoría `property-get`. **VOLUMEN MASIVO**, este filtro debe ser exacto |
| `property-set` (.Foo = value) | 5,250 | 1,527 | 6,777 | Asignación a propiedad — categoría `property-set` |
| `bang-get` (!Foo) | 12 | 202 | 214 | Bang operator — categoría `bang-get` |
| `bang-set` (!Foo = value) | 0 | 10 | 10 | Bang assignment — categoría `bang-set` |
| `do-fields-collection` (.Fields(...)) | 609 | 167 | 776 | DAO.Recordset.Fields() — **`dao-field-get` candidate** |
| `me-property` (Me.X) | 7,384 | 0 | 7,384 | Forms / self-reference — categoría `property-get` con receiver Me |
| `with-member` (.X al inicio de línea) | 2,122 | 745 | 2,867 | Inside `With` block — categoría `member-with` |
| `call-keyword` (`Call Foo`) | 51 | 118 | 169 | Statement-form antiguo — categoría `call` |
| `application-run` (Application.Run) | 0 | 10 | 10 | Late-bound dinámico — filtrar como runtime |
| `callbyname` (CallByName) | 0 | 2 | 2 | Late-bound dinámico — filtrar como runtime |
| `doevents` (VBA.DoEvents) | 1,282 | 61 | 1,343 | Runtime API — **filtrar como runtime en TODOS los rounds** |

## Conteo de receivers (prefix `Foo.`)

| Receiver | .cls | .bas | Significado |
|---|---|---|---|
| `Me.` | 7,432 | 13 | Forms / self. Refleja peso del layer UI. |
| `DoCmd.` | 1,176 | 18 | Access form-API. Runtime — filtrar. |
| `VBA.` | 1,282 | 216 | DoEvents + runtime. Filtrar. |
| `Scripting.` (FSO) | 429 | 649 | Late-bound. Filtrar. |
| `Forms(` | 30 | 32 | Access form-opening. Cross-form — qualify. |
| `Application.` | 55 | 83 | Application-level. Filtrar. |
| `Debug.` | 17 | 110 | Debug.Print. Filtrar. |
| `DAO.` | 77 | 919 | DAO API — `dao-field-get` o `dao-property` según shape. |
| `Access.` | 0 | 4 | Runtime. Filtrar. |
| `Collection.` | 0 | 0 | No encontrado — `Collection.Add` aparece como `bang`/`property-set` sin prefijo. |

## Hallazgos críticos para los rounds

### 1. `property-get` es 41k hits — el filtro de round-3 debe clasificarlo bien O MUERE

Si round-3 emite `property-get` como tipo pero deja items de `property-get` con `metadata.stub=true` (los que el parser no logra enlazar), el filtro consumer-side tendrá que filtrar **41k filas por la fuerza**. Esa es exactamente la situación que rompió #114.

Round-3 debe:
1. `property-get` cuando es lectura (`.Foo` sin `=` ni `(` después)
2. `property-set` cuando es asignación (`.Foo =`)
3. `dao-field-get` cuando es `rs.Fields("X")` o `rs!X`
4. **`NUNCA`** `property-get` para un synthetic stub cuyo target NO existe como property declaration

### 2. Conteo de `property-get` no debe contar comentarios ni string literals

El script excluye líneas que empiezan con `'` (comentario VBA) y `Attribute` (declaration headers). Si round-3 los cuenta, la cardinalidad reportada будет artificially baja para `property-get` lectura vs set.

### 3. `do-events-collection` (776 hits) — `dao-field-get` debería incluir tanto `rs!X` como `rs.Fields("X")`

Los bang-get (214) **bajo** porque la mayoría del acceso DAO va por `.Fields("X")`. Si round-3 clasifica solo bang-get como `dao-field-get`, deja 776 misses. Round-3 debe tener `dao-field-get` para ambos shapes.

### 4. `Application.Run` + `CallByName` — dinámico puro

12 hits en total. Round-3 podría filtrarlos como `runtime-filtered` sin clasificación de kind — son dynamic dispatch, no son calls declarables. Si los clasifica como `call`, aparecen como falsos missing.

### 5. `with-member` — implica `bar` dentro de `With foo` debe ser `foo.bar`

2122 hits en .cls — peso significativo en forms. Round-3 (y round-5 post-extraction resolver) deben resolver `.X` dentro de With contra el receiver del With. Sin eso, los edges de `with-member` quedan stub=true.

### 6. `me-property` — formas casi exclusivamente acceden a sí mismas

7384 hits en .cls集中在 forms. `Me.X` es legítimamente una `property-get` con receiver Me. Pero **NO** debe marcarse como `qualified-call` ni como runtime — Me.X es siempre lectura/escritura de self-control. Round-3 puede darlo como `property-get` directamente sin classification extra (receptor conocido).

### 7. Ejemplos concretos (1 por patrón, archivo:línea:match)

```
src/forms/Form_Form0BDOpciones.cls:12   —  Foo(arg)              — "EstablecerModeloCadenaJerarquica("
src/forms/Form_Form0BDOpciones.cls:55   —  Me.X                 — "Me.EntrarComo"
src/forms/Form_Form0BDOpciones.cls:358  —  .Foo = value         — ".Enabled ="
src/forms/Form_FormIndicador.cls:133    —  !Field               — "!RiesgosIdentificados"
src/forms/Form_FormAnexos.cls:65        —  .X (inside With)     — ".Title"
src/forms/Form_Form0BDOpciones.cls:50   —  VBA.DoEvents         — "VBA.DoEvents"
src/forms/Form_FormGestionRiesgosDatosGenerales.cls:79  —  Call Foo  — "Call "
src/classes/Anexo.cls:370               —  .Fields("X")         — ".Fields("
src/modules/Constructor.bas:63         —  .OpenRecordset       — ".OpenRecordset("
src/modules/Constructor.bas:69         —  .MoveFirst           — ".MoveFirst" (with-member)
src/modules/MigracionCorreosEnviadosIDEdicionLong.bas:54  —  Application.Run  — late-bound
src/modules/Test_CalidadCacheRiesgo.bas:637               —  CallByName        — late-bound
```

## Recomendaciones para el maintainer

Para round-3 (clasificación `unresolved_refs.reference_kind`):

1. **Mantener los 9 kinds** propuestos en el prompt (`call`, `qualified-call`, `property-get`, `property-set`, `bang-get`, `bang-set`, `dao-field-get`, `dao-field-set`, `unqualified-ident`, `member-with`).
2. **Añadir `runtime-filtered` como categoría adicional** para `VBA.DoEvents` y `Application.Run` y `CallByName` — no es `call`, es runtime. Si no se filtran, contaminan el reporte.
3. **Tratar `Me.X` como `property-get` antes de clasificar como `qualified-call`** — el receptor es conocido.
4. **Documentar `do-fields-collection` vs `!Field` como ambos en `dao-field-get`** — la mayoría del código DAO va por `.Fields("X")`.

Para round-5 (post-extraction resolver):

1. **Considerar `property-get` stubs además de los stubs de `paren-call-bare`**. Los 41k hits de property-get son mayoría — si round-5 solo arregla paren-call-bare, deja 95% del valor sin tocar.
2. **Resolver `.X` dentro de With contra el receptor del With** — 2867 hits en with-member. Sin esto, edges sintéticos.

## Script reproducible

```bash
# Desde la raiz del repo gestion_riesgos:
node tools/scratch-vba-syntax-profile.mjs 2>&1 | tee /tmp/vba-syntax.md
```

Re-ejecutable en cualquier codebase VBA. El script es standalone, no tiene deps más allá de Node 22+.

## Limitaciones

- El grep **cuenta todas las ocurrencias en el texto**, no desambigua entre uso real y decorativo. La cardinalidad aquí es **límite superior**; falsos positivos en cosas como comments (ya filtrados) y strings literales son posibles.
- El conteo **no incluye el contenido de string literals**. Ej. `Call("Foo")` dentro de `Debug.Print` no es llamado real.
- El script **no desambigua entre dos `.cls` que tengan el mismo nombre en distintos subdirs**. No es un problema en este repo pero podría serlo en otros.

Refs: gestion_riesgos#69 + ardelperal/codegraph-vba#108 (round-3) + ardelperal/codegraph-vba#110 (round-5).
