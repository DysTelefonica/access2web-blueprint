# Inventario de requisitos para Calidad

> **Issue:** #1494

## Para qué sirve

Antes de migrar a web, Calidad decide qué hace hoy cada aplicación de la Lanzadera y qué debe seguir haciendo. El inventario tiene dos lectores:

1. **Calidad**, que no es personal técnico, recibe un Excel. Cada fila es una funcionalidad o una regla de negocio, con el formulario y el botón tal como los ve en pantalla, y una columna para decidir.
2. **El backlog web** usa los mismos identificadores. Cada funcionalidad que se mantiene apunta a su evidencia en el código y a la ficha del PRD (`docs/03-aplicaciones/<app>/prd/`).

## Fuente y generación

La fuente versionada es `docs/03-aplicaciones/<app>/requisitos.yaml`, una por aplicación. El Excel no se versiona: se genera con

```bash
python scripts/build_requirements_xlsx.py docs/03-aplicaciones/<app>/requisitos.yaml -o <salida>.xlsx
```

El generador valida el fichero y falla con un mensaje claro si falta un campo obligatorio, si un identificador se repite o si un formulario de `formularios` no queda cubierto ni justificado.

## Esquema de `requisitos.yaml`

```yaml
aplicacion: Gestión de Riesgos          # nombre que conoce el usuario
prefijo: RIE                            # prefijo de los identificadores
repositorio: DysTelefonica/GESTION_RIESGOS
rama: release/2026-008.2.rc1
commit: 99ffa81df3f56a80e6f713e3afc0fb73fd9b3b25

formularios:                            # todos los src/forms/*.cls del commit
  - interno: Form_FormEdicion
    titulo: Edición del proyecto         # Caption visible; vacío si no tiene
    estado: cubierto                    # cubierto | no aplicable
    motivo: ""                          # obligatorio si es no aplicable

features:
  - id: RIE-F001
    area: Ediciones                     # agrupación de negocio (menú, módulo)
    nombre: Publicar edición            # verbo + objeto, lenguaje de negocio
    descripcion: >-                     # una o dos frases para Calidad
      Cierra la edición en curso y la deja publicada para consulta.
    formulario: Edición del proyecto    # título visible
    control: Publicar edición           # texto visible del botón o campo; vacío si no aplica
    tipo: botón                         # botón | campo | lista | menú | automático | informe
    roles: [Calidad, Técnico]           # quién puede usarla; [] si cualquiera
    resultado: >-                       # qué deja hecho, en lenguaje de negocio
      La edición queda publicada y se envía un aviso a los responsables.
    tecnico:                            # capa de backlog; no se muestra a Calidad
      formulario: Form_FormEdicion
      control: ComandoPublicar
      evidencia: [src/forms/Form_FormEdicion.cls:120]
      prd: prd/formularios/Form_FormEdicion.md
    reglas:
      - id: RIE-F001-R01
        descripcion: Solo se puede publicar si todos los riesgos tienen responsable.
        evidencia: [src/classes/Edicion.cls:340]
```

## Reglas de redacción

- `nombre`, `descripcion`, `resultado` y las reglas se escriben para alguien que usa la aplicación pero no programa. Ni nombres de tablas, ni de clases, ni SQL.
- Una regla es una condición que tiene que cumplirse, una validación, un cálculo o un efecto colateral (un correo, un documento generado, un cambio de estado). Cada regla es una sola afirmación comprobable.
- Toda feature y toda regla citan su evidencia `ruta:línea` en el commit fijado. Sin evidencia no entra.
- Se describe lo que hace hoy el código, no lo que debería hacer.
- Los procesos sin botón (tareas al abrir, envíos automáticos, cálculos al guardar) también son features, con `tipo: automático`.
- Castellano peninsular formal.

## Convenciones validadas en el piloto

El piloto de Gestión de Riesgos (área de publicación) fijó estas convenciones; todas las aplicaciones las siguen.

- **Título del formulario.** Muchas ventanas sustituyen su `Caption` en tiempo de ejecución (por ejemplo por «Usuario conectado: …»), y el `Caption` guardado puede contener nombres de personas. En esos casos `formulario` es el rótulo de cabecera visible, con las partes variables como «…», o el texto de la pestaña cuando el subformulario se muestra dentro de una. Nunca se copian nombres de personas.
- **Reglas compartidas.** Si varias funcionalidades aplican el mismo conjunto de comprobaciones, se describen una vez en la funcionalidad principal y las demás tienen una regla que remite a ellas por su identificador.
- **Posibles defectos.** Se describe el comportamiento actual y se añade un comentario YAML `# posible defecto: …` en la regla. No se corrige ni se interpreta.
- **Formularios no aplicables.** Solo con motivo real y comprobado: nadie los abre, son subformularios sin comportamiento propio, o son copias de prueba.

## Excel

Una hoja `Requisitos` con una fila por feature y, debajo, una fila por cada una de sus reglas:

| ID | Área | Funcionalidad / regla | Descripción | Formulario | Botón o campo | Quién | Resultado | Decisión | Comentario |
|---|---|---|---|---|---|---|---|---|---|

La columna Decisión ofrece Mantener / Retirar en las filas de feature y Mantener / Cambiar / Quitar en las de regla. Una hoja `Formularios` lista la cobertura, y una hoja `Origen` da el repositorio, la rama, el commit y los totales.
