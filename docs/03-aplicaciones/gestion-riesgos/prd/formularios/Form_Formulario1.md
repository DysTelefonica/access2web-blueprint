# Form_Formulario1

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_Formulario1.cls` (15 líneas) y `src/forms/Formulario1.form.txt` · **PRD:** [README](../README.md)

Formulario **huérfano**. Su único código es una línea de depuración; el resto son controles que pertenecen a otra aplicación. No hace nada. Es uno de los dos formularios **añadidos** respecto al snapshot: está en `main` y no en staging.

> Lo que sigue describe el objeto tal como está en el commit fijado, **no** una capacidad de la aplicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| **Nadie** | Búsqueda por nombre en `src`: no hay ningún `DoCmd.OpenForm` ni referencia que lo abra | búsqueda en los 195 ficheros de producción, 0 resultados fuera de sus propios ficheros |

No acepta parámetros, no abre otros formularios y no declara eventos hacia un contenedor.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Escribe el color de una etiqueta en la ventana de depuración | `:13-15` |

El diseño tiene nueve controles: `ComandoAceptar`, `ComandoCambiarTipoAMoC`, `ComandoEliminarAMoC`, `ComandoRechazar`, `Detalle` y `lbl1Complementado` (`.form.txt`). **Ninguno tiene handler**: el `.cls` no declara ni un solo `Sub` de evento.

## Flujos

### Apertura (`Form_Load`, `:13-15`)

`Debug.Print Me.lbl1Complementado.ForeColor`. Es una lectura de una propiedad de presentación con el único fin de inspeccionarla durante el desarrollo. No altera nada y no se ve en ejecución.

No hay flujo normal, alternativo ni de error: **no hay control de errores, no hay entrada y no hay salida**.

## Reglas, validaciones y cálculos

Ninguna. El formulario no implementa reglas.

## Lecturas y escrituras

Ninguna. No toca la base de datos ni el disco.

## Coexistencia

Sin implicaciones. Un objeto que no se abre no migra: la decisión es **retirarlo**, y conviene hacerlo de forma explícita para que no reaparezca en el alcance de la migración.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** un formulario sin código funcional, sin llamadores y con controles de otra aplicación.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-15.** Los nombres de los controles **no corresponden a esta aplicación**: `ComandoCambiarTipoAMoC` y `ComandoEliminarAMoC` hablan de cambio de tipo a Acuerdo Marco o Contrato, que es un concepto de Expedientes, no de Gestión de Riesgos. Es un resto de otro proyecto, copiado tal cual.
- **H-GR-16.** `ComandoAceptar`, `ComandoRechazar` y `ComandoEliminarAMoC` son **botones que no hacen nada**: el `.cls` no declara ningún manejador. Si alguien llegara a abrirlo por el diseñador de Access, tendría botones muertos a la vista.
- **H-GR-17.** `Form_Load` usa `Debug.Print` sin `On Error`: si `lbl1Complementado` no existiera en el binario (es decir, si el diseño fuera distinto del `.form.txt`), el formulario no abriría. Es poco relevante mientras esté huérfano, pero es una dependencia de diseño no declarada.
- **H-GR-18.** Está **en `main` pero no en el snapshot de staging**, al contrario que la mayoría de deltas de esta aplicación (que van de staging a `main`). Es el objeto más reciente de los tres que hay en producción y no llegaron al snapshot, junto con `EdicionSuministrador` y `ProyectoSuministrador`. No se puede fechar a partir del código.

**Propuestas web pendientes de decisión:** ninguna aprobada. La candidata es una: **retirar el objeto** del binario y del código exportado, y no darle ficha en la migración. Se deja constancia en la ficha para que la decisión sea explícita y no por omisión.

## Criterios de aceptación

Propuestos, no aprobados:

1. Retirado el formulario, ningún camino del código o del menú lo referencia.

Ese es el único criterio que se propone. No tiene comportamiento que preservar.

## Pendientes

- Confirmar en el binario que no se abre desde una macro, desde la cinta de opciones ni desde el formulario de inicio (código no exportado a `src`).
- Si el objeto se retira, decidir si también se retira del origen; es una decisión del propietario sobre su repositorio.