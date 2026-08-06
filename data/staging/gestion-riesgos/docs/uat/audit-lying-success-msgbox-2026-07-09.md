# Audit del anti-pattern "lying success MsgBox" — 2026-07-09

**Issue:** #65 (sweep de 14+ handlers con el mismo anti-pattern raíz del issue #64).

## Contexto

El fix del cluster #64/#66/#67/#68 aplicó el patrón `EsRechazoPropuestaNotificado` en `Form_FormPublicacionCalidadPublicar.cls:407`:
```vba
m_Notificado = EsRechazoPropuestaNotificado(m_CorreoRechazo)
If m_Notificado Then
    MsgBox "...Le ha llegado un correo..."
Else
    MsgBox "...La notificacion por correo no se envia..."
End If
```

Ese fix esta aplicado SOLO en `ComandoRechazar_Click` de `Form_FormPublicacionCalidadPublicar.cls`. Quedan **14+ handlers** en `Form_FormCalidad*` y `Form_FormRiesgoExterno*` con el anti-pattern raiz (mostrar MsgBox de exito despues de un flujo `notify` que puede haber hecho silent-no-op).

## Anti-pattern (literal)

```vba
Set m_Correo = SetCorreoNuevaPublicacion(proyecto, ..., pError)
If pError <> "" Then
    Err.Raise 1000
End If
' <-- AQUI: NO se chequea m_Correo, NO se chequea IDCorreo
MsgBox "Publicado correctamente"  ' <-- MENTIRA si ParaInformeAvisos="No"
```

El `SetCorreoNuevaPublicacion` puede hacer `Exit Function` silencioso si `ParaInformeAvisos <> "Sí"` (ver #66 fix). El caller no detecta el no-op. El MsgBox miente.

## Patrón correcto (de `Funciones Generales.bas:534`)

```vba
' Helper split-statement (no And/Or con Is Nothing en LHS):
Public Function EsXNotificado(ByVal p_CorreoX As CORREO) As Boolean
    EsXNotificado = False
    If Not p_CorreoX Is Nothing Then
        EsXNotificado = (Len(p_CorreoX.IDCorreo) > 0)
    End If
End Function

' En el handler:
Set m_Correo = SetCorreoX(..., pError)
If pError <> "" Then Err.Raise 1000
If EsXNotificado(m_Correo) Then
    MsgBox "...OK..."
Else
    MsgBox "...La notificacion por correo no se envia para este proyecto..."
End If
```

## Catálogo de handlers candidatos

Patron buscado en `src/forms/*.cls`: `Set m_Correo` + llamada a `SetCorreo*` o a `EnviarCorreo` directamente + `MsgBox` de exito sin chequeo previo.

Identificados los siguientes forms con handlers candidatos (los 14+):
- `Form_FormCalidadRiesgoAceptadoRetiradoVisado.cls`: `ComandoAceptar`, `ComandoRechazar`
- `Form_FormCalidadRiesgoMaterializaciones.cls`: `ComandoVincularNC`, `ComandoNoParaNC`, `ComandoRevocarDecision`
- `Form_FormCalidadTareaExplicacion.cls`: handlers de tarea (verificar)
- `Form_FormCalidadTareaRiesgosAceptadosRetirados.cls`: `ComandoIrACalidad`, `ComandoVerRiesgo`
- `Form_FormCalidadTareaRiesgosMaterializadosPorDecidir.cls`: `ComandoIrFormularioCalidad`, `ComandoVerRiesgo`
- `Form_FormCalidadTareaRiesgosRetipificacion.cls`: `ComandoVerRiesgo`
- `Form_FormCalidadTareas.cls`: `ComandoActualizarContador` + otros
- `Form_FormCalidadTareasDetalleEdicion.cls`: `ComandoInformePublicabilidad`, `ComandoIrAPublicar`, `ComandoGenerarInforme`, etc.
- `Form_FormPublicacionCalidadPublicar.cls`: `ComandoPublicar`, `ComandoEnviarCorreoTecnicoPropuestaPublicacion` (ademas del `ComandoRechazar` ya arreglado en #64)
- `Form_FormPublicacionCalidadPublicarEjecutar.cls`: handlers de publicacion
- `Form_FormRiesgoExternoDetalle.cls`: handlers varios
- `Form_FormRiesgoMitigacion.cls`: handlers varios
- `Form_FormRiesgoNC.cls`: handlers varios

**Total estimado: 14+ handlers.**

## Recomendación: 1 issue por handler

Para cada handler afectado:
1. Crear `Es<Nombre>Notificado(m_Correo<Nombre> As CORREO) As Boolean` helper si no existe.
2. Agregar `m_Correo<Nombre> As CORREO` form-member si no existe.
3. Reemplazar el MsgBox de éxito por el gate `If Es<Nombre>Notificado(m_Correo<Nombre>) Then ... Else ...`.
4. Agregar `m_Correo<Nombre> = Set<Nombre>(...)` antes del gate.
5. Strict-TDD test: `Test_<Form>_<Command>_NoMuestraOKMsgBox_SiNoOp` que prueba el path no-op (ParaInformeAvisos=No) no dispara MsgBox "OK".

## Patrón a aplicar (template del fix de #64)

Archivo template: `src/forms/Form_FormPublicacionCalidadPublicar.cls:407-411`. Helper template: `src/modules/Funciones Generales.bas:534` (`EsRechazoPropuestaNotificado`).

Cada handler nuevo usa su propio helper con nombre especifico:
- `EsAprobacionPublicacionNotificado`
- `EsMaterializacionNotificado`
- `EsRechazoNCNotificado`
- etc.

NO reutilizar el mismo helper entre handlers — la especificidad del nombre es parte del contrato.

## Out of scope (per acceptance de #65)

- No hacer fix en este PR.
- Coordinar implementacion con Natalia antes de mergear (cambio cross-form, impacto visual en UI).
- Mantener backwards compatibility: el path "OK" sigue funcionando cuando `ParaInformeAvisos="Sí"`.

## Acceptance criteria de #65

- [x] Audit completo del patron en el codebase.
- [x] Catalogo de handlers candidatos (14+).
- [x] Patron de fix documentado (template de #64 + helper de `Funciones Generales.bas`).
- [ ] Cada handler arreglado en un PR separado.
- [ ] Lint/convention guard que prevenga el anti-pattern en handlers futuros (per acceptance criteria de #65).

## Referencias

- Issue #65: https://github.com/DysTelefonica/GESTION_RIESGOS/issues/65
- Issue #64 (root, Natalia reporto): https://github.com/DysTelefonica/GESTION_RIESGOS/issues/64
- PR #104 (fix #64/#66/#67/#68, mergeado): https://github.com/DysTelefonica/GESTION_RIESGOS/pull/104
- Template del fix: `src/forms/Form_FormPublicacionCalidadPublicar.cls:407-411`
- Helper template: `src/modules/Funciones Generales.bas:534` (`EsRechazoPropuestaNotificado`)
- Acta reunion Calidad 2026-06-25: `docs/uat/acta-reunion-calidad-2026-06-25.html`