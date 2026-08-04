# 02 · Topología del ecosistema — Expedientes (ciclo de vida)

## Propósito

Documenta el ciclo de vida del expediente como agregado de negocio transversal: alta, estados, entidades relacionadas, suministradores, responsables, cierre/baja y consultas. Es la referencia para entender la dependencia que casi todas las demás aplicaciones tienen respecto a `IDExpediente`, `CodExp`, `CodExpLargo` y `Nemotecnico`.

## Qué va aquí

- Modelo del expediente: `TbExpedientes`, entidades asociadas, suministradores, responsables.
- Estados y transiciones conocidas, con su evidencia documental o de código.
- Reglas de unicidad y equivalencias entre identificadores (cuando se verifiquen).
- Visión de consumidores actuales (sin inferir los no demostrados).

## Estado del contenido

**Placeholder.** No contiene hallazgos propios; se poblará durante el Lote 2 de `exploration.md`.

## Fuentes de autoridad

1. Documentación en `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES`.
2. Código en `C:\00repos\codigo\00_EXPEDIENTES\00_main`.
3. Dysflow en modo solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Las equivalencias entre identificadores requieren tabla de correspondencias con cardinalidad.
- Las excepciones por aplicación (por ejemplo, No Conformidades con `Nemotecnico`) se documentan aparte y no se mezclan con la regla general.
- Los avisos de configuración (`path-mismatch` en `.dysflow/project.json`) se registran sin corregir en esta fase.

## Checklist

- [ ] Cada estado del expediente lleva su origen (código, doc, prueba).
- [ ] Los consumidores se listan solo cuando la evidencia los sostiene.
- [ ] Las excepciones están marcadas como `Excepción` con su aplicación y flujo.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/identificadores-y-glosario.md` al iniciar el Lote 2 para fijar el contrato de identificadores.
