# 02 · Topología del ecosistema — Lanzadera (identidad y permisos)

## Propósito

Documenta el rol transversal de Lanzadera como origen de identidad, permisos y catálogo de aplicaciones. Sirve de referencia para todas las fichas de `03-aplicaciones/` y para `04-integraciones-y-operacion/`.

## Qué va aquí

- Hechos verificados sobre `TbAplicaciones`, `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`.
- Patrón de arranque, comando VBA y aperturas desde otras aplicaciones.
- Contrato de identidad: `UsuarioRed`, `CorreoUsuario`, `VBA.Command`, clases `UsuarioAplicacionPermisos`.
- Tabla de aplicaciones vigentes del catálogo (cuando se extraiga del Lote 1).

## Estado del contenido

**Placeholder.** Aún no se ha ejecutado el Lote 1. No contiene hallazgos; solo describe qué se documentará y con qué reglas.

## Fuentes de autoridad

1. Documentación en `C:\00repos\documentacion\OPENSPEC\00_LANZADERA`.
2. Código en `C:\00repos\codigo\00_LANZADERA\00_main`.
3. Dysflow en modo solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Toda fila del catálogo debe llevar fuente, fecha y, si aplica, evidencia Dysflow.
- Los roles y permisos se citan con su tabla origen; no se infieren desde el nombre del rol.
- Las excepciones documentadas (por ejemplo, cuentas técnicas) se etiquetan como tales.

## Checklist

- [ ] Cada afirmación lleva etiqueta `Verified-*` / `Intended` / `Likely` / `Divergent`.
- [ ] Las referencias a tablas y procedimientos citan módulo y nombre exacto.
- [ ] No se mezclan permisos de entornos (producción, staging, sandbox) sin marcarlo.

## Siguiente paso

Poblar este documento al ejecutar el Lote 1 de `exploration.md`; detener y solicitar aprobación antes de pasar al Lote 2.
