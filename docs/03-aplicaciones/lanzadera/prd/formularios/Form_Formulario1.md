# Form_Formulario1

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_Formulario1.cls` (6 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de pruebas con el nombre por defecto de Access: tres botones sin código y un rectángulo cuyo color imprime en la ventana de depuración al abrirse. No tiene origen de datos ni lee o escribe en la base de datos. Las citas son `Form_Formulario1.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| — | Ninguna referencia en `src`: ni `OpenForm`, ni `SourceObject`, ni nombre en módulos o clases | búsqueda de `Formulario1` en `src/**/*.cls` y `*.bas` sin resultados |

No abre nada.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | `Debug.Print` del color de fondo de `Cuadro0` | `:4-6`, `.form.txt:30` |
| `ComandoGrabar` («Grabar») | botón | sin evento | Sin comportamiento | `.form.txt:214-216` |
| `ComandoRiesgoEliminar` («Eliminar», ayuda «Eliminar Riesgo») | botón | sin evento | Sin comportamiento | `.form.txt:258-260` |
| `ComandoSHPS` | botón con imagen | sin evento | Sin comportamiento | `.form.txt:320-323` |
| `Cuadro0` | rectángulo | — | Solo se lee su `BackColor` | `:5`, `.form.txt:302` |

## Flujos

- **Apertura:** imprime el color de `Cuadro0` en la ventana inmediata (`:5`); no hay más efectos.
- **Errores:** sin manejador; no hay operaciones que puedan fallar salvo la lectura del control (`:4-6`).

## Reglas, validaciones y cálculos

Ninguna.

## Lecturas y escrituras

Ninguna.

## Coexistencia

Sin impacto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Objeto de pruebas huérfano: ningún código de `src` lo abre. El rótulo «Eliminar Riesgo» (`.form.txt:260`) sugiere que se copió de otra aplicación.
- `Debug.Print` como única lógica (`:5`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no migrarlo.

## Criterios de aceptación

Ninguno: no aporta capacidad de negocio.

## Pendientes

- Confirmar que ningún menú en datos ni macro del binario lo abre; el nombre aparece en `Lanzadera.accdb`, pero la extracción de solo lectura de menús y macros no se ha hecho.
- Decidir en el inventario si se clasifica como **no aplicable**.
