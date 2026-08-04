# 04 · Integraciones y operación — Identidad, arranque y permisos

## Propósito

Recoge los patrones compartidos de arranque, autenticación y resolución de permisos: `VBA.Command`, clases `UsuarioAplicacionPermisos`, llamadas `getdbLanzadera` y `getdbExpedientes` y consumo de `TbUsuariosAplicacionesPermisos`.

## Qué va aquí

- Secuencia de arranque típica por aplicación (sin inventar la de cada una).
- Resolución de permisos: tabla origen, clase, helper y errores conocidos.
- Apertura de aplicaciones desde otras (`VBA.Command` y similares).

## Estado del contenido

**Placeholder.** Se poblará con los hallazgos de los Lotes 1 a 8 y se consolidará en el Lote 9.

## Fuentes de autoridad

1. `C:\00repos\codigo\<app>\00_main` por aplicación.
2. Inspección Dysflow solo lectura.
3. CodeGraph y CodeGraph-VBA para trazado de símbolos.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Los permisos se documentan por tabla origen; no por nombre de rol.
- El patrón `getdbLanzadera` / `getdbExpedientes` se describe como familia, no como implementación única.
- Las cuentas técnicas y excepciones se etiquetan como tales.

## Checklist

- [ ] Cada patrón lleva módulo, clase o helper que lo implementa.
- [ ] Los roles se referencian al catálogo de Lanzadera, no se duplican aquí.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/lanzadera-identidad-permisos.md` para evitar duplicación.
