# 07 · Migración — Modelo de dominio agnóstico

## Propósito

Empieza a abstraer el modelo de dominio del ecosistema legacy a términos independientes de la plataforma. Sirve de entrada para `07-migracion/matriz-legacy-a-web.md` y para futuros `proposal/spec/design/tasks` de la fase SDD de migración. **No** selecciona stack ni define arquitectura objetivo.

## Qué va aquí

- Entidades de dominio candidatas: `Usuario`, `Aplicacion`, `Permiso`, `Expediente`, `Solicitud`, `Riesgo`, `NoConformidad`, `Suministrador`, `SLA`, etc., con su definición provisional.
- Atributos e identificadores transversalizados (`IDExpediente`, `CodExp`, `Nemotecnico`) como contrato de equivalencias.
- Reglas de negocio invariantes, independientes de Access.
- Excepciones a la dependencia del expediente (con su flujo, no como regla general).

## Estado del contenido

**Placeholder.** No se ha poblado todavía; la primera versión se redactará al cierre del Lote 9 de `exploration.md`.

## Fuentes de autoridad

1. `05-capacidades/` y `03-aplicaciones/`.
2. `02-topologia-ecosistema/`.
3. `C:\00repos\codigo\<app>\00_main` y `C:\00repos\documentacion`.
4. Dysflow solo lectura, como validación puntual.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- Una entidad se modela solo si aparece en al menos una capacidad con confianza `Verified-static` o superior.
- Los identificadores se tratan como contrato de equivalencias; no se normalizan a un único valor en esta fase.
- Las reglas de negocio se redactan en lenguaje de dominio, sin referencias a Access, DAO o SQL.

## Checklist

- [ ] Cada entidad cita las capacidades que la sostienen.
- [ ] Las equivalencias entre identificadores están en su propia tabla.
- [ ] Ningún modelo introduce todavía un framework o servicio web concreto.

## Siguiente paso

Mantener este modelo sincronizado con cada ficha de capacidad que se cierre; reabrir si una capacidad rompe una invariante.
