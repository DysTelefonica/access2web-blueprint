<!--
Documento de capacidad CAP-006 - Comercial (CRUD generico de entidades).
Linaje: PRUEBA-003 REFAC-2a (single-record forms) + REFAC-2b (gestion forms) + dysflow_get_schema de TbComerciales + dysflow_test_vba del Helper_EntidadCRUD.
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Capacidad: Comercial (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-006
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Responsable / autoridad de producto**: Pendiente de confirmacion
- **Ultima verificacion**: 2026-06-16 mediante `dysflow.get_schema` (TbComerciales) + `dysflow.test_vba` (9/9 verde slice-refac-2a/2b) + `dysflow.compile_vba` (PASS)
- **Confianza global**: alta - ver 7 (BR-006-01..03 `Verified-runtime`; BR-006-04..05 `Verified-static` - el happy path del helper requiere MockEntidad.cls)

## 1 Intencion de negocio (~ proposal SDD) - POR QUE
- **Proposito**: permitir al administrador dar de alta, modificar y eliminar Comerciales que se asocian a expedientes via `TbExpedientesComerciales`. Los Comerciales son listas planas de proveedores externos (no se mezclan con TbSuministradores, que es la lista de contratistas).
- **Usuarios / perfiles**: administrador (solo el admin puede crear/editar/eliminar; los gestores solo consultan).
- **Problema que resuelve**: el CRUD estaba duplicado en 2 forms (~350 LOC cada uno):
  - `Form_FormComercial.cls` (single-record) — con `ComandoRegistrar_Click`, `HaHabidoCambios`, `EstablecerDatos` (~200 LOC)
  - `Form_FormComercialesGestion.cls` (lista) — con `ComandoAlta/Editar/Eliminar/Buscar_Click` (~250 LOC)
  - Misma logica repetida 10 veces (Comercial, CPV, Ejercito, Suministrador, Lugar, Oficina, Organo, PECAL, RAC, Grado) — 2 forms cada uno = 20 forms total, ~5000 LOC duplicadas.
- **Valor de negocio**: el helper extrae la logica generica (copy form->obj, compare obj vs form, eliminar via class). Las 10 entidades usan el mismo helper con distintos field names. Reduccion de ~5000 LOC duplicadas a ~500 LOC en el helper + 10x 5 LOC de rewire = ~550 LOC.
- **No-objetivos**: no es soft delete / undo; no maneja auditoria de quien-edito-que; no sincroniza con sistemas externos de proveedores.
- **Origen de la intencion**: spec PRUEBA-002 3.2 (BR-006-01..02) + PRUEBA-003 6.1 (helper stateless) + refactor de 10 entities con el mismo patron (slice REFAC-2a/2b).
- **Referencia de tracker de origen**: PRUEBA-003 REFAC-2a (single-record forms, commit `dc04d1d` + 5 commits de rewire) + REFAC-2b (gestion forms, commit `a104f5b` + 1 commit de rewire).

## 2 Contrato de comportamiento (~ spec SDD) - QUE ? ANCLA DE REGRESION

### Escenarios (Dado / Cuando / Entonces)
- **DADO** el usuario hace click en "Alta" en `Form_FormComercialesGestion` **CUANDO** completa los campos `Comercial` y `DESCRIPCION` y hace click en "Registrar" **ENTONCES** `m_ComercialOp.Registrar m_ComercialAlInicio, m_Error` ejecuta el INSERT/UPDATE en `TbComerciales` y dispara el evento `Alta` o `Editado` del form hijo.
- **DADO** el usuario selecciona un Comercial de la lista **CUANDO** modifica `Comercial` o `DESCRIPCION` y hace click en "Registrar" **ENTONCES** `Helper_EntidadCRUD.HaHabidoCambiosGenerico(m_ComercialAlInicio, m_ValoresForm, Array("Comercial","DESCRIPCION"), p_Error)` retorna `True` y el flujo permite guardar.
- **DADO** el usuario selecciona un Comercial de la lista **CUANDO** hace click en "Eliminar" y confirma **ENTONCES** `Helper_EntidadCRUD.EliminarEntidadGenerico(m_ComercialOp, "Comercial", m_ComercialSeleccionado, "...", m_Error)` llama `m_ComercialOp.Eliminar` via `CallByName`, dispara el evento `Eliminar`, y refresca la lista.
- **DADO** `p_Obj=Nothing` o `p_Entidad=Nothing` **CUANDO** se invoca el helper **ENTONCES** retorna error con p_Error poblada (sin crashear).

### Reglas de negocio

| ID regla | Enunciado | Autoridad | Aplicada en codigo? | Prueba | Confianza |
|---|---|---|---|---|---|
| **BR-006-01** | `Helper_EntidadCRUD.CopiarCamposAObjeto` con `p_Obj=Nothing` retorna error sin crashear | spec PRUEBA-003 6.1 (input validation) | Si - `Helper_EntidadCRUD.CopiarCamposAObjeto` (early return) | `Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError` - PASA 2026-06-15 | Verified-runtime |
| **BR-006-02** | `Helper_EntidadCRUD.HaHabidoCambiosGenerico` con `p_ObjInicial=Nothing` retorna `True` (es un alta) | spec | Si - `Helper_EntidadCRUD.HaHabidoCambiosGenerico` (early return) | `Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue` - PASA 2026-06-15 | Verified-runtime |
| **BR-006-03** | `Helper_EntidadCRUD.EliminarEntidadGenerico` con `p_Operaciones=Nothing` o `p_Entidad=Nothing` retorna error | spec | Si - `Helper_EntidadCRUD.EliminarEntidadGenerico` (early return) | `Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError` + `_EntidadNothing_PueblaError` - PASA 2026-06-16 | Verified-runtime |
| **BR-006-04** | `Helper_EntidadCRUD.CopiarCamposAObjeto` copia los 2 fields (`Comercial`, `DESCRIPCION`) de `m_ValoresForm` a `m_ObjComercialActivo` via `CallByName` | spec | Si - rewire de `Form_FormComercial.ComandoRegistrar_Click` (commit `770ca37`) | (no testeable sin MockComercial.cls; impl coverage via compile_vba) | Verified-static |
| **BR-006-05** | `Form_FormComercialesGestion.ComandoEliminar_Click` elimina el Comercial seleccionado via el helper | spec | Si - rewire (commit `a104f5b`) | (no testeable sin abrir UI; impl coverage via compile_vba) | Verified-static |

### Mensajes de error
- `MSG-006-01`: `"CopiarCamposAObjeto: p_Obj is Nothing"` (input validation).
- `MSG-006-02`: `"CopiarCamposAObjeto: p_Valores is Nothing"` (input validation).
- `MSG-006-03`: `"CopiarCamposAObjeto: campo '<nombre>' no esta en p_Valores"` (campo faltante).
- `MSG-006-04`: `"CopiarCamposAObjeto: <Err.Description> (campo: <nombre>)"` (DAO/late-binding error).
- `MSG-006-05`: `"HaHabidoCambiosGenerico: <Err.Description> (campo: <nombre>)"` (DAO/late-binding error).
- `MSG-006-06`: `"EliminarEntidadGenerico: p_Operaciones is Nothing"` (input validation).
- `MSG-006-07`: `"EliminarEntidadGenerico: p_Entidad is Nothing"` (input validation).
- `MSG-006-08`: `"EliminarEntidadGenerico: <Err.Description>"` (DAO/late-binding error).

(Los mensajes estan en espanol porque la audiencia son administradores que ven el UI; las claves internas son los IDs de regla.)

## 3 Especificacion tecnica (~ design SDD) - COMO
- **Modulo del helper**: `src/modules/Helper_EntidadCRUD.bas` (stateless, `.bas`, ~190 LOC con las 3 funciones + helpers).
- **Modulos de los forms**: `src/forms/Form_FormComercial.cls` (single-record) + `src/forms/Form_FormComercialesGestion.cls` (lista).
- **DAO consumido**: el helper es generico via `CallByName` (late binding), no toca DAO directamente. Los forms usan `getdb()` (de `Variables Globales.bas`) via `m_ComercialOp.Eliminar` (class method que hereda de `XxxOperaciones`).
- **Sin popup, sin MsgBox**: el helper NO muestra UI. El `MsgBox` de confirmacion esta DENTRO del helper `EliminarEntidadGenerico` (parametro `p_Confirmacion`).
- **Convenciones respetadas**:
  - `Optional ByRef p_Error As String` ultimo parametro (per convencion EXPEDIENTES).
  - `Optional ByVal p_Db As DAO.Database = Nothing` (en `CopiarCamposAObjeto` y `HaHabidoCambiosGenerico`, aunque via `ResolveDb`).
  - Late binding via `CallByName` para evitar acoplamiento con entity classes especificas.
- **Decisiones de diseno**:
  - **Generico via Dictionary + CallByName**: el helper recibe un `Scripting.Dictionary` de field name -> value. No conoce las entity classes (Comercial, CPV, etc.) - solo escribe via `CallByName p_Obj, m_Nombre, VbLet, m_Valor`. Asi, el mismo helper sirve para las 10 entidades.
  - **`EliminarEntidadGenerico` con `CallByName` para operaciones**: el helper no necesita saber la entity class. Recibe `p_Operaciones` (cualquier `XxxOperaciones`), `p_NombrePropEntidad` (el nombre de la property que recibe la entidad: "Comercial", "CPV", etc.), y `p_Entidad`. Hace `CallByName p_Operaciones, p_NombrePropEntidad, VbSet, p_Entidad` y luego `CallByName p_Operaciones, "Eliminar", VbMethod, p_Error`.
  - **No shim en FUNCIONES UTILES**: el helper es llamado directamente desde el .cls del form (`Helper_EntidadCRUD.CopiarCamposAObjeto ...`), no a traves de un shim. Asi evitamos el problema de "ambiguous name" (visto en REFAC-1b con `GenerarConsultaExpedientes`).
  - **Limitacion de tests (reconocida)**: el happy path no es testeable sin `MockEntidad.cls` (una clase con properties reales como `Comercial` y `DESCRIPCION`). Los tests actuales cubren error paths (input validation, mock invalido).

## 4 Operacion (~ tasks SDD) - QUIEN/HACE CUANDO
- **PR owner**: PR-REFAC-2a/2b (cerrado en este slice).
- **Reviewers**: 1 (mantenedor).
- **Tests anadidos en este slice**: 9 verde (6 del helper + 3 de eliminacion).
- **Tareas realizadas**:
  - T1a.3 (helper) creado: `Helper_EntidadCRUD.bas` con `CopiarCamposAObjeto` + `HaHabidoCambiosGenerico` + `ResolveDb` (commit `dc04d1d`).
  - T1a.4 (tests) verdes: 6 tests error paths (commit `dc04d1d`).
  - T1a.5 (binary) sync: `Expedientes.accdb` (commit `6532f62`).
  - T2a.6 (rewire 11 single-record forms): CPV, Comercial, Ejercito, Lugar, Modificado, Oficina, Organo, PECAL, RAC, Grado, Suministrador (commits `0f7b1b0` + `770ca37` + `7d814d2` + `a99d9e4`).
  - T2b.1 (helper extension): `EliminarEntidadGenerico` agregado (commit `a104f5b`).
  - T2b.6 (rewire 10 gestion forms): Comerciales, CPVs, Ejercitos, Grados, Lugar, Oficinas, Organo, PECALES, RACS, Suministradores (commits `a104f5b` + `b59b03f`).
  - T2b.7 (binary) sync: `Expedientes.accdb` (despues de cada commit de rewire).

## 5 Referencias cruzadas
- Spec original: `staging-alignment-prueba-002/specs/capabilities/CAP-006-comercial.spec.md` (BR-006-01..02 wish).
- Spec del helper: `staging-alignment-prueba-003/specs/helpers/Helper_EntidadCRUD.spec.md` (TBD).
- Apply progress: `staging-alignment-prueba-003/apply-progress.md` (PR-REFAC-2a/2b sections).
- Tasks: `staging-alignment-prueba-003/tasks.md` (PR-REFAC-2a/2b sections).
- Feature doc: `docs/features/F-EXPORT-helper-expedientes-export-excel.md` (estructura similar; este CAP es el companion).
- Forms involved: `Form_FormComercial.cls` (single-record) + `Form_FormComercialesGestion.cls` (lista).
- Helper: `src/modules/Helper_EntidadCRUD.bas`.
- Entities que siguen el mismo patron (9 mas, todas rewireadas): CPV (CAP-007), Ejercito (CAP-008), Suministrador (CAP-009), LugarEjecucion (CAP-010), PECAL (CAP-011), RAC (CAP-012), GradoClasificacion (CAP-013), OrganoContratacion (CAP-014), OficinaPrograma (CAP-015).
- Engram: `sdd/expedientes/capability-map/explore` (mapa de capabilities).

## 6 Lagunas explicitas
- **Happy path del helper NO testeable** (visto en `Test_Helper_EntidadCRUD_*`): el mock Dictionary no tiene las properties esperadas (Comercial, DESCRIPCION), asi que `CallByName` falla antes de poder testear el happy path. Para cerrar la brecha haria falta:
  - `MockEntidad.cls` — una clase con properties Public (Comercial As String, DESCRIPCION As String) que el helper pueda escribir/leer via `CallByName`.
  - Tests del happy path para las 3 funciones del helper.
  - Estimado: 1 commit con el .cls + ~6 tests nuevos.
- **`ComandoEliminar_Click` con cancelacion del usuario**: el helper retorna "" si el usuario cancela, pero el form actualmente NO chequea este caso (siempre hace `RaiseEvent Eliminar` y `ComandoBuscar_Click`). Refinar el rewire para solo continuar si `m_Error = ""`.
- **EstablecerDatos (entity -> form)**: el form rewireado solo cubre ComandoRegistrar y HaHabidoCambios. El `EstablecerDatos` (3er With block en cada form) sigue inline. Para extraer: agregar `Helper_EntidadCRUD.CopiarObjetoAForm(p_Obj, p_Form, p_NombresCampos) As String`. Trivial extension, no scope de este slice.
- **32 tests VS→VR pendientes**: 9 entities x 1-2 reglas base cada una + otras 23 reglas. Estos tests verificarian el CRUD funcionando end-to-end (alta -> consulta -> edicion -> eliminacion). Target: PR-D1/D2.

## 7 Resumen de cobertura
- Total reglas: 5 (BR-006-01..05).
- `Verified-runtime`: 3 (BR-006-01, BR-006-02, BR-006-03) — error paths del helper.
- `Verified-static`: 2 (BR-006-04, BR-006-05) — el rewire del form compila, pero el happy path no se testea (gaps documentados).
- Tasa de cobertura verde: 9/9 tests slice-refac-2a/2b (3+ entidades ademas de Comercial).
- Coverage delta desde PRUEBA-003 inicio: 0% -> 60% (3/5 reglas) del helper cubierto, 0% -> 100% (2/2 forms) rewireados y compilando.
- Deuda: BR-006-04/05 happy paths, EstablecerDatos no rewireado, cancelacion del usuario no diferenciada, 32 tests VS→VR globales.
