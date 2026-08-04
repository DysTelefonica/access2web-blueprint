# 03 · HPS_Solicitudes

## Propósito

Carpeta de descubrimiento de **HPS_Solicitudes**: sistema Access documentado de solicitudes, con discrepancia crítica — no se ha localizado un checkout principal bajo `C:\00repos\codigo` con ese nombre.

## Estado

- **Documentación:** localizada y rica (PRD 01–05, ERD, cambios).
- **Checkout principal:** **no localizado** como `00_HPS_SOLICITUDES` bajo `C:\00repos\codigo`.
- **Dysflow:** no aplica todavía (sin target).
- **Discrepancia:** la documentación describe un sistema Access, pero no existe checkout principal con ese nombre en el directorio inspeccionado.

## Lote asociado

Lote 8 de `exploration.md` — HPS_Solicitudes (último, tras resolver la identidad).

## Entregable previsto (condicionado)

1. Confirmar si HPS_Solicitudes es aplicación independiente, línea histórica o parte de HPS.
2. Si es independiente: checkout, binarios, Dysflow y capacidades (solicitudes, workflow, correo, automatizaciones).
3. Si está integrada en HPS: documentar la relación y reasignar la documentación.
4. Dependencias con Lanzadera (usuarios/permisos) y Expedientes (`TbSolicitudes.IDExpediente`).

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` (PRDs 01–05, ERD y cambios).
2. `C:\00repos\codigo` — barrido para identificar checkout equivalente o renombrado.
3. Inspección Dysflow solo lectura (una vez resuelto el target).
4. Engram solo como contexto histórico.

## Reglas de evidencia

- **No se infiere** que HPS_Solicitudes sea HPS ni viceversa hasta resolver la discrepancia.
- La dependencia documental con `TbExpedientes.IDExpediente` se documenta como `Fuerte documental` y se marca como `pendiente de runtime`.
- Mientras no haya binario localizado, no se extraen capacidades runtime.

## Checklist

- [ ] Pregunta 1 de `08-decisiones-y-preguntas-abiertas.md` resuelta con el usuario.
- [ ] Checkout/binary localizado (o se documenta formalmente la inexistencia).
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Detener el descubrimiento de HPS_Solicitudes hasta que el usuario aclare la identidad del repositorio/binario.
