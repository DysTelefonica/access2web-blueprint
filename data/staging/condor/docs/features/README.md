# Funcionalidades técnicas

Esta carpeta queda reservada para documentación de funcionalidades técnicas o subflujos que soportan capacidades de negocio.

La navegación principal debe empezar por [`../capabilities/capabilities-index.md`](../capabilities/capabilities-index.md). Una funcionalidad técnica solo debe documentarse aquí cuando aporte detalle reutilizable que no pertenezca a una única capacidad.

## Convenciones

- Documentar primero la capacidad de negocio; después enlazar funcionalidades técnicas si hacen falta.
- Mantener los hechos de comportamiento con el mismo registro de confianza: `Verified-runtime`, `Verified-static`, `Intended`, `Likely`, `Divergent`.
- No marcar pruebas como actuales si no se ejecutaron mediante Dysflow.
- Las funcionalidades técnicas de pruebas deben seguir `access-vba-tdd` v2.4.2: retorno JSON desde `Public Function`, fixture sandbox propio, schema-first, inyección explícita de `DAO.Database`, cardinalidad en mutaciones, manifests Dysflow con procedimiento global único sin calificación, sin UI/`Debug.Print` y sin escribir en `TbConfiguracionBackends`.
- Si una funcionalidad técnica permite probar formularios, debe favorecer extracción a helper, servicio o ViewModel; la lógica de negocio no debe quedar encerrada en eventos de formulario.

## Funcionalidades técnicas candidatas

| Funcionalidad técnica | Capacidades relacionadas | Estado |
|---|---|---|
| Backend seleccionable y `getdb()` único | Todas | Pendiente |
| Generación Word/PDF y mapeos de plantilla | PC, CDCA, CDCASUB, PCSUB | Pendiente |
| Gestión de adjuntos y cierre formalización | PC, CDCA, CDCASUB, PCSUB | Pendiente |
| Workflow y precondiciones por estado | Todas las solicitudes | Pendiente |
| Fixtures sandbox y manifest de pruebas Dysflow | PCSUB y futuras capacidades documentadas | Pendiente |
