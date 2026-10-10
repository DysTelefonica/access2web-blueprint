---
name: ai-slop-discipline
description: "Trigger: cualquier trabajo asistido por IA — escribir o revisar PRs, planificar fixes, redactar issues, modificar código, prosa o tests — y antes de empezar CUALQUIER cambio. Comprobación obligatoria de 23 puntos."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.0"
  last_verified: 2026-10-07
  scope: ['universal']
  auto_invoke: ['trabajo asistido por IA', 'empezar cualquier cambio', 'revisar un PR']
  tiers: ['universal']
---

# Disciplina anti-slop para trabajo asistido por IA

## El principio

La persona que presenta el cambio es responsable de TODA decisión de diseño
y de todo lo que entrega: seguridad, corrección, mantenimiento, defensa en
revisión. La IA es herramienta de trabajo mecánico —sintaxis, refactors,
andamiajes de tests—, no autora de diseño.

Prueba concreta: elija una línea al azar del diff y pregúntese: «¿Puedo
defender esta línea como mi propia decisión?». Si no puede, la línea se
retira; no se retira la auditoría.

## Declaración y atribución

- **Declaración**: la ayuda material de IA en código, tests, documentación,
  diseño o revisión sustantiva se declara en el cuerpo del PR indicando
  herramienta o modelo (si se conoce), alcance de la ayuda y verificación
  realizada. Formato, correcciones ortográficas, autocompletado trivial y
  navegación no se declaran.
- **Atribución**: los commits no llevan crédito a la IA
  (`Co-Authored-By`, `Reviewed-by`, `Tested-by`, `Signed-off-by` ni
  equivalentes). El revisor puede pedir explicación, procedencia o tests
  adicionales, y rechazar lo que el autor no sepa explicar ni defender.
- **No engañar**: nada de afirmaciones sin verificar, resultados que no
  ocurrieron, APIs, rutas, evidencias o resultados de test inventados.

## Comprobación antes de empezar cualquier trabajo (23 puntos)

Responda cada punto antes de escribir. Si alguna respuesta es «no» o «no
sé», PARE y pregunte.

### Alcance y autoría

1. **Alcance.** ¿Lo que voy a hacer está en los criterios de aceptación de
   la issue? Si no, es fuera de alcance.
2. **Autoría.** ¿El humano podría defender cada línea como decisión suya?
   Si no, no la escribo.
3. **Duplicado.** ¿Ya existe algo que lo haga (función, detector, parser,
   regla, sección)? Si existe, lo reutilizo y no creo una segunda fuente de
   verdad —un segundo detector de la misma causa es deuda, no cobertura—.
4. **Fuente de verdad.** ¿Decido a partir del dato real o de una
   heurística? Nada de detectar por nombre de fichero, por subcadena o por
   texto libre cuando existe el dato estructurado que responde.
5. **Fallo cerrado.** Si mi gate o control no puede medir, ¿falla o pasa en
   verde? Nada de verde sin medición: truncados, activación por
   autoselección —un gate que solo rige si el propio consumer se declara—
   y valores desconocidos tratados como correctos.
6. **No engañable.** ¿Se puede satisfacer mi gate con prosa o con un
   ejemplo inventado sin cumplir la regla? Un test que afirma sin probar
   no es cobertura.
7. **Esquivar un gate.** ¿Estoy cambiando contenido para que un gate deje
   de quejarse? Entonces el defecto está en el gate: se corrige el gate,
   no la prosa.
8. **Tests reales.** ¿Cada test sale de un caso de la issue o de un fallo
   real, no de un ejemplo que invento yo?
9. **Afirmaciones verificables.** ¿Cada «funciona», «está en verde» o «no
   recursa» lo he EJECUTADO y medido? Si no, no lo afirmo.
10. **Contexto ligero.** ¿Estoy metiendo en un fichero que lee la IA
    (AGENTS, bloques propagados) contenido que se podría cargar bajo
    demanda? Las reglas completas van en la skill; el índice solo despacha.

### Cobertura exacta de la issue

11. **Ni más ni menos que la issue.** Haz la lista de criterios de
    aceptación. Cada cambio corresponde a un criterio. Si un criterio no
    tiene cambio, falta trabajo. Si un cambio no tiene criterio, sobra: se
    quita o se pregunta.
12. **Pruebas exactas.** Cada criterio tiene al menos una prueba que falla
    sin el cambio (RED observado) y pasa con él. Ningún test cubre algo que
    la issue no pide, y no hay dos tests que comprueben lo mismo. Cada test
    se nombra por el criterio que demuestra.
13. **Tabla de trazabilidad en el cuerpo del PR.** `criterio → fichero(s)
    cambiado(s) → test que lo demuestra`: una fila por criterio, ninguna
    fila sin test, ningún cambio fuera de la tabla. Si un criterio queda
    pendiente, se declara explícitamente y el PR lleva enlace, no cierre.

### Sin metalenguaje ni relleno

14. **Sin metalenguaje en los artefactos.** El código, los comentarios, los
    documentos y los mensajes de commit describen el sistema, no el
    proceso. Nada de «como IA…», «en este PR he…», «nota honesta» ni
    narrar cómo se llegó al cambio. Eso va, si acaso, en la conversación o
    en el cuerpo del PR, nunca en el código.
15. **Comentarios solo con información que el código no da:** el porqué,
    una restricción o un riesgo. Nada de comentarios que repitan lo que
    hace la línea, historiales ni referencias a la conversación.
16. **Sin relleno ni adjetivos de marketing:** robusto, inteligente,
    completo, perfecto, sin fisuras, de forma elegante. Si una propiedad
    importa, se demuestra con un test, no con un adjetivo.
17. **Sin generalidad especulativa:** nada de parámetros, opciones,
    abstracciones, capas ni «por si acaso» que la issue no pide. Tres
    líneas repetidas son mejores que una abstracción prematura.

### Sin nombres ni rutas personales

18. **Sin nombres ni rutas personales.** ¿Cito personas, logins,
    `C:\Users\…`, `/home/…` o `personal/<usuario>/` en algo que se propaga
    o se documenta? Se usan roles («cualquier mantenedor») y rutas
    instaladas (`.agents/skills/<skill>/…`).

### Historia, procesos y estilo

19. **Historia intacta.** ¿Voy a reescribir algo publicado (rebase,
    `--amend`, force-push)? Se actualiza con `git merge origin/main`.
20. **Procesos seguros.** ¿Lanzo subprocesos que puedan recursar o
    colgarse? Todo `subprocess.run` lleva `timeout=`, y las pruebas con
    riesgo van en un scope acotado.
21. **Sin marcadores vacíos:** nada de TODO, FIXME, «pendiente de
    implementar», funciones vacías ni tests saltados. Lo que no se hace se
    declara como pendiente en la issue o en el PR.
22. **Estilo del entorno:** el código y los documentos siguen el idioma,
    el tono y las convenciones de lo que ya hay alrededor. Sin emojis ni
    formato llamativo donde el resto no los usa.
23. **Sin reescribir lo que no se toca:** nada de reformatear, renombrar
    ni reordenar fuera del alcance para que «quede mejor». El diff contiene
    solo lo necesario.

## Las tres firmas del slop (detectar antes de enviar)

1. **Prosa paralela que duplica una autoridad existente.** Una sección,
   fichero o contrato nuevo que reafirma lo que otro ya dice. Un ejemplo
   real: dos secciones del mismo contrato redactando la misma regla de
   validación con distintas palabras; la revisión la pilló, pero ya era deuda.
2. **Afirmaciones amplias sin contrato testeable.** «automático»,
   «adaptativo», «a prueba de fallos», «elegante»: si no admite el test
   concreto más pequeño que la demuestre, es marketing, no contrato. O se
   escribe el test o se quita la afirmación.
3. **Tests que verifican lo que la IA inventó, no casos reales.** Un test
   que exige strings de ejemplo que nadie escribiría, un montaje
   contrived o un escenario que no aparece en los criterios de la issue.
   El peor tipo de slop: sobrevive a la revisión y se vuelve contrato
   portante.

## Autorrevisión previa al diff

Antes de mostrar el diff a la persona:

- Por cada fichero o sección nuevo: ¿ya existe uno que lo cubra?
- Por cada afirmación: ¿puedo escribir ahora su test concreto?
- Por cada test: ¿de qué fallo real o criterio de la issue nace?
- Por cada línea fuera del alcance de la issue: ¿por qué está aquí?

Si alguna respuesta es «no lo sé», esa línea es candidata a retirarse.

Y para la persona, la prueba de defensa: elija una línea al azar y
responda si puede defenderla como decisión propia. Si no, se retira.

## Límite de delegación a subagentes

Toda tarea delegada lleva un límite de superficie explícito:

```
Solo puede modificar:
- [lista de ficheros/secciones/funciones permitidos]

No puede:
- Tocar ningún otro fichero
- Reescribir secciones ajenas
- Añadir comportamiento que la issue no pide
- Añadir tests de ejemplos inventados
- Crear contratos, enums o tipos no requeridos
```

El fallo que evita: una delegación de «amend las cláusulas existentes»
interpretada como licencia para reescribir dos secciones enteras más.

## Causa, invariante y tamaño del cambio

Antes de proponer un fix: identifique la causa subyacente y la invariante
responsable, y sepa defender por qué el cambio es proporcional. Prefiera el
cambio más pequeño que restaure esa invariante, sin autoridad duplicada,
abstracciones innecesarias ni complejidad no relacionada.

Inaceptable: enviar output no revisado; enmascarar el síntoma o desplazar
el fallo dejando la invariante rota; cambios globales fuera del alcance;
delegar en el revisor la comprensión, la validación o la reparación.

## Prohibido

- Inventar una sección nueva «por completitud» cuando una cláusula
  existente ya lo cubre.
- Cambios «ya que estamos» ajenos a la issue.
- Strings de ejemplo en la prosa que no se puedan trazar a una entrada
  real.
- Defender el slop con «pero así queda mejor»: si no se pidió, no se añade.
- Tratar sugerencias de revisores automáticos como decisiones de diseño:
  son input, no autoridad.
