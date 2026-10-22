# Form_FormEnConstruccion

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormEnConstruccion.cls` (3 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla estática de «en construcción»: un único control de imagen sin código de comportamiento. No tiene origen de datos, eventos ni lecturas o escrituras. Las citas son `Form_FormEnConstruccion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| — | Ninguna referencia en `src`: ni `OpenForm`, ni `SourceObject`, ni nombre en módulos o clases | búsqueda de `FormEnConstruccion` en `src/**/*.cls` y `*.bas` sin resultados |

No abre nada.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Imagen38` | imagen | — | Muestra la imagen `EnConstruccion.jpg` incrustada | `.form.txt:245-250` |

El formulario no declara eventos (`OnLoad`, `OnOpen` u otros) en el `.form.txt`. El volumen del `.form.txt` (unas 24 000 líneas) corresponde a los datos binarios de la imagen incrustada.

## Flujos

Ninguno. El módulo solo declara `m_Error` y no lo usa (`:3`).

## Reglas, validaciones y cálculos

Ninguna.

## Lecturas y escrituras

Ninguna.

## Coexistencia

Sin impacto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** formulario de una sola vista (`ViewsAllowed = 1`), sin bordes ni botones de navegación (`.form.txt:5-15`), que muestra una imagen.

**Posibles defectos (sin corregir, a verificar):**

- Objeto aparentemente huérfano: ningún código de `src` lo abre ni lo aloja. Podría abrirse desde un menú guardado en tabla o en el binario.
- Variable `m_Error` declarada sin uso (`:3`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no migrarlo; si hace falta un aviso de funcionalidad no disponible, resolverlo con un componente genérico de la web.

## Criterios de aceptación

Ninguno mientras no se confirme que se usa. Si se decide conservar el aviso: una opción de menú sin implementación muestra un aviso de «en construcción» sin errores.

## Pendientes

- Confirmar si algún menú en datos (tablas de menú, macros o `Lanzadera.accdb`) lo abre; el nombre aparece en el binario, pero la extracción de solo lectura de menús y macros no se ha hecho.
- Decidir en el inventario si se clasifica como **no aplicable**.
