# Cutover Pack — Expedientes

Pack de cutover para la migración de Expedientes. Define el UAT, cutover y retirada verificable.

**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-9)
**Spec**: `uat-cutover-legacy-retirement.md` (EXP-CAP-057..063)

## UAT (User Acceptance Testing)

**Alcance**: `EXP-CAP-001..063` por perfil, estados vacíos, errores y concurrencia.

**Perfiles**:
- Administrador
- Gestor de área
- Responsable calidad
- Consulta (solo lectura)

**Estados vacíos**: Verificar que los estados vacíos se manejan correctamente (sin datos, sin errores).

**Errores**: Verificar que los errores se manejan correctamente (sin crash, sin datos corruptos).

**Concurrencia**: Verificar que la concurrencia se maneja correctamente (sin race conditions, sin datos duplicados).

**Criterio de UAT**: Nunca compara pantallas. Verifica comportamiento, no paridad visual (D-EXP-5).

## Cutover

**Procedimiento**:
1. **Ensayo**: Ejecutar el cutover en un entorno de staging para verificar el procedimiento.
2. **Delta final**: Ejecutar el extractor para capturar los cambios finales del legacy.
3. **Freeze**: Congelar el legacy para evitar cambios durante el cutover.
4. **Reconciliación**: Ejecutar la reconciliación para verificar la integridad de la migración.
5. **UAT**: Ejecutar el UAT para verificar el comportamiento del nuevo sistema.
6. **Switch**: Cambiar el routing para dirigir el tráfico al nuevo sistema.

**Criterio de cutover**: El cutover es reversible antes del switch. Después del switch, el rollback requiere restaurar desde backup.

**Rollback**: Si el cutover falla, ejecutar el procedimiento de rollback del reconciliation runbook.

## Retirada verificable

**Alcance**: `EXP-CAP-057..063` (retirada de Win32, procesos Win32, OLE/ActiveX, globals, selector backend, popups, menú JSON-hub).

**Procedimiento**:
1. **Verificar la retirada**: Confirmar que cada componente legacy fue retirado correctamente.
2. **Documentar la retirada**: Registrar la retirada en el log de migración.
3. **Verificar el nuevo sistema**: Confirmar que el nuevo sistema está operativo y los datos están intactos.

**Criterio de retirada**: La retirada es verificable. Cada componente legacy tiene un criterio de retirada claro y verificable.

**Retirada de Win32 red**: Verificar que las llamadas Win32 a la red fueron retiradas correctamente.

**Retirada de procesos Win32**: Verificar que los procesos Win32 fueron retirados correctamente.

**Retirada de OLE/ActiveX**: Verificar que los controles OLE/ActiveX fueron retirados correctamente.

**Retirada de globals**: Verificar que las variables globales fueron retiradas correctamente.

**Retirada de selector backend**: Verificar que el selector de backend fue retirado correctamente.

**Retirada de popups**: Verificar que los popups fueron retirados correctamente.

**Retirada de menú JSON-hub**: Verificar que el menú JSON-hub fue retirado correctamente.

## Verificación

- [ ] El UAT fue ejecutado y aprobado
- [ ] El cutover fue ejecutado y verificado
- [ ] La reconciliación fue ejecutada y aprobada
- [ ] El switch fue ejecutado y verificado
- [ ] La retirada fue verificada y documentada
- [ ] El rollback fue documentado y probado
