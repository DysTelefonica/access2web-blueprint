# 07 · Migración — Matriz legacy → web

## Propósito

Matriz viva que asocia capacidades legacy con futuros componentes web. **No** define el stack objetivo ni la arquitectura destino: solo deja trazada la correspondencia para que una fase SDD posterior pueda descomponerla en `proposal/spec/design/tasks`.

## Qué va aquí

- Una fila por capacidad: `capacidad legacy · componente web candidato · riesgo principal · dependencias · excepciones · confianza`.
- Identificación de capacidades con dependencia crítica de UI Access (lógica a extraer a servicio/use-case).
- Identificación de capacidades candidatas a conservar como reporte o solo-lectura en web.
- Trazabilidad a la ficha de capacidad correspondiente en `05-capacidades/`.

## Estado del contenido

**Placeholder.** La primera versión se redactará al cierre del Lote 9 de `exploration.md` y se refinará al iterar sobre las fichas de capacidad.

## Fuentes de autoridad

1. `05-capacidades/`.
2. `07-migracion/modelo-dominio-agnostico.md`.
3. `03-aplicaciones/<app>/` por si una capacidad es exclusiva de una aplicación.
4. Dysflow solo lectura, como validación puntual.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- Un componente web candidato se nombra solo si hay capacidad legacy que lo justifique.
- Las capacidades con lógica en evento de formulario se marcan como `lógica a extraer` y se referencian a su módulo/handler.
- Las excepciones (`Nemotecnico` en No Conformidades, etc.) viajan con su fila, no se difuminan.

## Checklist

- [ ] Cada fila cita la ficha de capacidad y la aplicación que la sostiene.
- [ ] Las capacidades marcadas como `lógica a extraer` enlazan con `06-seguridad-y-trazabilidad.md` si tocan permisos.
- [ ] Ningún componente web introduce todavía decisiones de stack.

## Siguiente paso

Cruzar con `08-decisiones-y-preguntas-abiertas.md` cuando aparezcan divergencias o preguntas de mayor valor.
