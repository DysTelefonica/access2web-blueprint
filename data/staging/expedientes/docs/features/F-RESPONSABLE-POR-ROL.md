<!--
Hoja de feature - modResponsablePorRolHelper / modResponsablePorRolService / modResponsableFiltroHelper.
Capa: docs/features/ (segundo nivel).
Issue: #71 (responsable-combos-e2e-tdd).
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Feature: Responsables por rol desde `TbResponsablesPorRol`

## 0 Identidad
- **ID de feature**: F-RESPONSABLE-POR-ROL
- **Issue**: #71 (responsable-combos-e2e-tdd)
- **Modulos**:
  - `src/modules/modResponsablePorRolHelper.bas` (DAO + whitelist cerrada)
  - `src/modules/modResponsablePorRolService.bas` (envelope JSON + AddItem a combos)
  - `src/modules/modResponsableFiltroHelper.bas` (helper puro string<->string)
  - `src/modules/Test_ResponsablePorRolHelper.bas` (8 atomos TDD)
  - `src/modules/Test_ResponsableFiltroHelper.bas` (9 atomos TDD)
- **Estado**: source listo (en `feature/71-responsable-combos-e2e-tdd`); binary import bloqueado por resolucion de worktree MCP (ver §8)
- **Source**: tests + production + form wiring
- **Confianza**: `Verified-static` hasta el compile humano + green atoms

## 1 Proposito y motivo
- **Proposito tecnico**: sustituir las listas hardcodeadas de responsables (Calidad/Seguridad) por una lectura desde `TbResponsablesPorRol` (backend) JOIN `TbUsuariosAplicaciones` (Lanzadera), con whitelist cerrada del rol, soporte para IDs "huerfanos" (usuario guardado en el expediente que ya no tiene el rol), y un service JSON que popula los combos del form.
- **Motivo de existencia**: `GuardadoAutomaticoHelper.bas:EstablecerCombos` hardcodeaba los responsables de Seguridad (lineas 1007-1011), y `Form_FormExpedientesGestion.cls` usa OptionGroups (`MarcoCalidad` 1-7, `MarcoRespSeguridad` 1-3) mapeados a nombres literales en `FUNCIONES UTILES.bas:getExpBusqueda` (lineas 1041-1067). Cualquier alta/baja de responsable requiere tocar codigo.
- **Decision de diseno clave**: separacion en 3 modulos con responsabilidad unica -- helper DAO, service JSON envelope, helper puro de traduccion string.  El filtro puro no toca SQL ni DAO; los tests del filtro no necesitan temp `.accdb`.  El helper DAO y el service se prueban con temp `.accdb` per-test (sandbox §5.5).

## 2 Contrato publico (firmas)

### `modResponsablePorRolHelper`
```vb
' Whitelist cerrada: "Calidad" | "Seguridad".  Otro valor -> Err.Raise 1000.
' p_db inyectado.  Nothing -> fallback a CurrentDb() del frontend.
' Devuelve Dictionary keyed por CStr(Id), value = Usuario.  Incluye ID
' "huerfano" pasado en p_IDResponsableActual aunque ya no tenga el rol
' (lookup defensivo en TbUsuariosAplicaciones).
Public Function GetUsuariosPorRol( _
    ByVal p_Rol As String, _
    ByVal p_IDResponsableActual As String, _
    Optional ByRef p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary
```

### `modResponsablePorRolService`
```vb
' Carga los combos IDResponsableCalidad / IDResponsableSeguridad de un
' form y devuelve el envelope JSON canonico {ok, value, payload, error,
' logs}.  Si frm no es Nothing, hace AddItem a los combos (Clear+AddItem
' para que sobrescriba cualquier lista previa).  Si p_db es Nothing, el
' helper usa CurrentDb() del frontend (que enrutara las linked tables a
' sus backends respectivos).  Los tests inyectan su propio p_db.
Public Function ExpedienteGeneral_EstablecerCombos( _
    ByRef frm As Form, _
    Optional ByVal p_OrphanCalidad As String = "", _
    Optional ByVal p_OrphanSeguridad As String = "", _
    Optional ByRef p_db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String
```

### `modResponsableFiltroHelper` (helper puro, sin SQL/DAO/Forms)
```vb
Public Const RESPONSABLE_FILTRO_TODOS  As String = "Todos"
Public Const RESPONSABLE_FILTRO_NA_ID   As String = "0"
Public Const RESPONSABLE_FILTRO_NA_LABEL As String = "N/A"

Public Function ResponsableFiltro_EsTodos(ByVal p_Valor As Variant) As Boolean
Public Function ResponsableFiltro_ValorAId(ByVal p_Valor As Variant, _
                                            Optional ByRef p_Error As String) As String
Public Function ResponsableFiltro_IdACmbRow(ByVal p_ID As String, _
                                             ByVal p_Dict As Scripting.Dictionary, _
                                             ByVal p_ConSentinel As Boolean) As String
Public Function ResponsableFiltro_ValorEsId(ByVal p_Valor As Variant, _
                                             ByVal p_ID As String) As Boolean
```

## 3 Form wiring (staging thin-form integration)
- **`Form_FormExpedientesGestion.cls:MarcoCalidad_AfterUpdate`** (linea ~3470): llama a `modResponsableFiltroHelper.ResponsableFiltro_EsTodos(CStr(Me.MarcoCalidad))` para detectar "sin filtro" via el helper puro en vez de hardcodear `Me.MarcoCalidad = 0`.  El OptionGroup sigue siendo la fuente de filtro actual; el call deja el wiring listo para la migracion a combos.
- **`Form_FormExpedientesGestion.cls:MarcoRespSeguridad_AfterUpdate`** (linea ~1966): idem para `Me.MarcoRespSeguridad`.
- **`Form_FormExpedienteGeneral.cls:EstablecerDatos`** (linea ~306): despues de la llamada legacy `EstablecerCombos frm:=Me, p_Error:=m_Error`, invoca `modResponsablePorRolService.ExpedienteGeneral_EstablecerCombos Me, "", "", Nothing, m_Error`.  El service hace Clear+AddItem sobre los combos Calidad/Seguridad, asi que sobrescribe los datos que la llamada legacy cargo desde cache.  La cache `m_ObjEntorno.ColUsuariosCalidad` sigue poblada por la llamada legacy, asi que `GuardadoAutomaticoHelper` que la lee no se rompe.

## 4 Schema-first fixture (regla fixture-first gate, docs/testing-fixture-first-gate.md)
- **Tablas tocadas**:
  - `TbResponsablesPorRol` (backend `Expedientes_datos.accdb`):
    - `IDResponsablePorRol`: `Long` (autonumber PK)
    - `IDUsuario`: `Long` (NOT NULL)
    - `Rol`: `Text(50)` (NOT NULL)
    - Indice unico `(IDUsuario, Rol)`
  - `TbUsuariosAplicaciones` (Lanzadera backend, linked):
    - `Id`: **`Integer` (16-bit, type=3, size=2)** -- este es el constraint que la primera version del fixture violaba con IDs `900001+`
    - `Nombre`, `UsuarioRed`, `CorreoUsuario`, `Activado`, `FechaAlta`, `FechaBaja`
- **DDL minimo en cada temp .accdb (mirror del schema real)**:
  - `TbResponsablesPorRol`: `IDResponsablePorRol COUNTER NOT NULL PRIMARY KEY` + `IDUsuario INTEGER NOT NULL` + `Rol TEXT(50) NOT NULL`
  - `TbUsuariosAplicaciones`: `Id SHORT NOT NULL` + resto (SHORT en DDL Access == dbInteger 16-bit)
  - Unique index `(IDUsuario, Rol)` en `TbResponsablesPorRol`
- **IDs de fixture (Integer-safe, deterministas)**:
  - Calidad: 30001..30005
  - Seguridad: 30006..30008
  - Cross-rol (Calidad + Seguridad): 30011
  - Baja logica (Calidad + FechaBaja): 30010
  - Orphan en `TbUsuariosAplicaciones` sin rol: 30012
- **Per-test temp `.accdb`** (sandbox §5.5): cada test crea su propio archivo en `C:\Users\adm1\AppData\Local\Temp\test-responsable-por-rol\<atom>_<timestamp>.accdb` con la DDL arriba.  El helper NUNCA toca el backend productivo en estos tests porque siempre recibe un `p_db` no Nothing (verificar que el helper respeta el db inyectado es parte del contrato).
- **Seed deterministic**: `InsertUsuario` y `InsertRol` insertan con IDs explícitos (no autonumericos aleatorios).  `InsertRol` deja que `COUNTER` asigne `IDResponsablePorRol` (no testeado en los asserts).

## 5 Bridge E2E <-> UAT
- **TDD manifest slice**: `tests/tests.vba.responsable-71.json` (18 atomos, todos `tags:["issue-71", ...]`)
- **Main manifest tag**: cada atomo aparece en `tests/tests.vba.json` con `tags:["issue-71", ...]` (mismos procedures + nombres)
- **UAT scenarios**: pendiente (ver §8 bloqueadores)

## 6 Reglas de diseno
1. **Whitelist cerrada** del rol en `ValidarRol`.  Cualquier valor fuera de {Calidad, Seguridad} -> `Err.Raise 1000`.  SQL nunca se interpola sin validar.
2. **Inyeccion explicita de `DAO.Database`** en todos los metodos DAO-touching.  Tests inyectan su temp `.accdb` (verificar el contrato de inyeccion es parte de los atoms).
3. **Convencion de error canonica**: `Optional ByRef p_Error As String` como ultimo parametro.  Quien falla setea `p_Error` y hace `Err.Raise 1000`.  Caller chequea `If p_Error <> "" Then Err.Raise 1000`.
4. **No `MsgBox`/`Debug.Print`** en test atoms (contract §1.1).
5. **Orden determinista** en el SQL: `ORDER BY U.Nombre` (alfabetico, case-insensitive en Jet/ACE default collation).
6. **Sin duplicados**: el `Dictionary` esta keyed por `CStr(Id)`; un usuario con multiples filas de rol colapsa a una entrada.
7. **Compatibilidad con huerfanos**: si `p_IDResponsableActual` es no-vacio y no figura entre los usuarios devueltos, se anade via lookup defensivo en `TbUsuariosAplicaciones`.  Asi un expediente antiguo no queda visualmente vacio.
8. **Service envelope canonico**: `{ok, value, payload, error, logs}` (runner contract §2).  Payload serializado via `JsonConverter.ConvertToJson`.
9. **UX sentinel Seguridad**: el combo Seguridad admite "sin responsable" como `0;N/A` (primer AddItem).  El helper de filtro reconoce este sentinel como equivalente a "sin filtro".

## 7 Refactor-safety / cambios inocuos
- Renombrar cualquier constante Public del helper puro no rompe nada (los tests las importan via su nombre canonico).
- Mover el lookup defensivo de `LookupUsuarioPorID` a un helper separado no cambia comportamiento.
- Cambiar `ORDER BY U.Nombre` por `ORDER BY U.Nombre, U.Id` agrega orden secundario sin alterar el set de usuarios.

## 8 Bloqueadores conocidos (al momento de redactar)
- **MCP write-block contra `feat71`'s binary**: el MCP `dysflow` activo esta resuelto al worktree `00_EXPEDIENTES_staging` (`get_capabilities.projectConfig.projectRoot == staging`).  Cualquier intento de importar al `Expedientes.accdb` de `feat71` devuelve `OUTSIDE_PROJECT_ROOT: Requested target ... is outside this worktree`.  Workaround: el humano debe re-arrancar el MCP desde `C:\00repos\codigo\00_EXPEDIENTES_feat71_responsable` y re-ejecutar el import de los 5 modulos en scope (`Test_ResponsableFiltroHelper`, `Test_ResponsablePorRolHelper`, `modResponsableFiltroHelper`, `modResponsablePorRolHelper`, `modResponsablePorRolService`).
- **TESTS NOT RUN** -- waiting for human compile (per instruccion explicita del orquestador).
- **Source vs binary drift residual**: el `Expedientes.accdb` en feat71 tiene una version "regresion 2026-07-14" de `modResponsablePorRolHelper.bas` que usa `getdb()` + `getdbLanzadera()` (DAO no permite JOIN entre dos backends Access).  El source en disco es la version pre-regresion (single JOIN via linked tables).  Esto es el drift `bothChanged` que el orquestador instruyo reconciliar source -> binary.  Si la regresion es deliberada, revertirla requiere re-mergear PR #70 (commit `9bf8f0d`).

## 9 Trazabilidad a UAT
- Card UAT usuario (pendiente): abrir FormExpedienteGeneral -> combos Calidad/Seguridad muestran usuarios reales desde TbResponsablesPorRol (no lista hardcodeada)
- Card UAT desarrollo: `Form_FormExpedientesGestion.cls:MarcoCalidad_AfterUpdate`/`MarcoRespSeguridad_AfterUpdate` invocan `modResponsableFiltroHelper.ResponsableFiltro_EsTodos` (visible en `Debug.Print` o via `Caller` introspection)