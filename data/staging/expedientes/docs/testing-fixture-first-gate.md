# Gate TDD fixture-first para Access/VBA

Este gate aplica a todos los tests Access/VBA de `PRUEBA-001` y a cualquier PR que se abra hacia `staging`.

## Regla central

Un test con datos solo es válido si prepara explícitamente los datos exactos que necesita en backend local/sandbox antes del Act.

No se acepta ningún test que pase porque el dato ya existía, porque encontró una fila parecida, o porque adaptó el assert al estado actual del backend.

## Precondición ERD/schema

Antes de escribir o aceptar cualquier fixture, hay que inspeccionar el schema real de todas las tablas tocadas.

La evidencia mínima por tabla es:

- Nombre de tabla.
- Primary key.
- Foreign keys y relaciones padre/hijo.
- Campos `Required` / `NOT NULL`.
- Tipo de cada campo poblado.
- Índices únicos o restricciones de unicidad relevantes.
- Dominios o valores válidos cuando apliquen.

Si falta esta información, el test queda bloqueado. No se adivinan campos opcionales, relaciones ni valores legales.

## Diseño obligatorio de fixture

Cada test o helper de fixture debe declarar:

1. **Objetivo del caso**: happy path, sad path o edge case.
2. **Tablas tocadas**: lectura y escritura.
3. **Grafo de fixture**: padres antes que hijos.
4. **Datos mínimos suficientes**: solo lo necesario para probar el comportamiento.
5. **Seed order**: orden exacto de inserción.
6. **Teardown order**: orden inverso, hijos antes que padres.
7. **Marcador determinista**: IDs reservados o token único del test.
8. **Aserciones fuertes**: estado concreto y efectos secundarios esperados.

## Prohibiciones

- `SELECT TOP 1` como preparación de fixture.
- Usar filas reales o datos de usuario como Arrange.
- Verificar que existe una fila y tratar eso como poblar.
- Conteos globales sin filtro de test.
- Snapshots arbitrarios como sustituto del seed.
- Ajustar el assert al dato encontrado.
- Compartir fixtures amplios si el test no verifica exactamente qué filas fueron insertadas.

Snapshot/restore solo puede usarse como protección de cleanup, nunca como estrategia de preparación.

## Casos exigidos

Cuando el comportamiento tenga ramas, los tests deben cubrir con fixtures propios:

- Camino feliz: datos válidos y completos.
- Camino triste: ausencia o invalidez controlada del dato requerido.
- Caso borde: duplicados, cardinalidad inesperada, valores límite o relaciones faltantes, según corresponda al contrato.

Cada caso debe poblar su propio estado ideal. No se reutiliza suerte ambiental.

## Evidencia requerida en revisión

Todo juez/reviewer de tests Access/VBA debe cargar `access-vba-tdd` y reportar:

- Skill cargada y ruta usada.
- Tests revisados.
- Tablas tocadas.
- Evidencia ERD/schema por tabla.
- Grafo de fixture.
- Seed order.
- Teardown order.
- Happy/sad/edge cases cubiertos.
- Dependencias restantes de datos preexistentes, si existen.
- Veredicto final: `CUMPLE` o `NO CUMPLE`.

## Criterio de salida para issue #16

La issue #16 solo puede darse por establecida cuando dos jueces ciegos revisan este gate y ambos dicen `CUMPLE` sin hallazgos críticos ni warnings reales.

Mientras un juez diga `NO CUMPLE`, no se abre PR hacia `staging` basado en esos tests.

## Prompt obligatorio para jueces

```text
Cargá y aplicá la skill Access/VBA TDD antes de revisar:
C:\Users\adm1\.config\opencode\skills\access-vba-tdd\SKILL.md

Objetivo: juzgar si los tests Access/VBA cumplen el gate fixture-first de PRUEBA-001.

Revisá contra estos criterios:
- Cada test con datos puebla explícitamente backend local/sandbox antes del Act.
- No depende de datos existentes, SELECT TOP 1, conteos globales, snapshots arbitrarios ni suerte ambiental.
- Antes de cualquier seed se conoce el ERD/schema real: PK, FKs, relaciones, campos Required/NOT NULL, tipos, índices/uniqueness y dominios.
- El fixture inserta padres antes que hijos y limpia hijos antes que padres.
- Cada caso tiene datos ideales propios para happy path, sad path o edge case.
- Las aserciones verifican comportamiento concreto y efectos secundarios sobre los datos sembrados.

Tu reporte debe incluir obligatoriamente:
- Tests revisados.
- Tablas tocadas.
- Evidencia ERD/schema consultada.
- Grafo de fixture.
- Seed order y teardown order.
- Dependencias de datos preexistentes detectadas.
- Veredicto: CUMPLE o NO CUMPLE.
```
