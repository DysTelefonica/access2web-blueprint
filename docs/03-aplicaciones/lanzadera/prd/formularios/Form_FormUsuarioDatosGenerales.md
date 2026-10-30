# Form_FormUsuarioDatosGenerales

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuarioDatosGenerales.cls` (393 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Datos Generales» de un usuario: alta y edición de la ficha en `TbUsuariosAplicaciones` (identidad, indicadores de administración, activación, caducidad de contraseña, responsable). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): carga y vuelca la global `m_ObjUsuarioActivo` (`Variables Globales.bas:100`) y persiste con `Usuario.Registrar`. Las citas son `Form_FormUsuarioDatosGenerales.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioGeneral` | `MENU2_Click` lo carga en `FormDetalle`, haya o no usuario activo | `Form_FormUsuarioGeneral.cls:77-81` |

Abre `FormUsuarioCambioPass` (cerrándolo antes si estaba abierto) desde `ComandoCambiarPass` (`:221-224`); ese formulario trabaja con el usuario conectado, no con el activo (`Form_FormUsuarioCambioPass.cls:82-87`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre`, `CorreoUsuario`, `Matricula`, `Telefono`, `Movil`, `UsuarioRed` | textos | — | Copia directa usuario ↔ control con `Nz` | `:116-130, 167-169, 261-265, 284`, `.form.txt:343, 616, 448, 504, 560, 1076` |
| `ComboEsAdministrador` («Admin») | combo `Sí;No` | — | Valor textual `EsAdministrador` | `:131, 266-270`, `.form.txt:677-679` |
| `ComboActivado`, `ComboUsuarioImborrable`, `ComboPasswordNuncaCaduca`, `ComboMantenerLanzaderaAbierta` | combos `Sí;No` | — | Se traducen a booleanos | `:133-157, 272-290, 301-307`, `.form.txt:741, 861, 1009, 1045` |
| `ComboPermisoPruebas` («Aplicaciones Prueba») | combo `Sí;No` | — | Vacío se muestra «No» | `:170-175, 291-295`, `.form.txt:1171-1173` |
| `FechaAlta`, `FechaBaja` | textos, fecha corta | — | Se muestran con `dd/mm/yyyy` | `:159-165, 282-283`, `.form.txt:799-800, 919-920` |
| `Responsable` | combo 2 columnas (ID oculto) | — | Todos los usuarios; se guarda el `UsuarioRed` del elegido | `:184-186, 296-300, 369-393`, `.form.txt:1355-1370` |
| `TextoTiempoParaExpirarPass` («Pass Caduca en») | texto | — | Texto calculado; no se guarda | `:187`, `classes/Usuario.cls:841-879` |
| `ID` | texto oculto y deshabilitado | — | ID del usuario | `:176, 318`, `.form.txt:1094-1108` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida cambios, vuelca y guarda | `:241-333`, `.form.txt:370-372` |
| `ComandoCambiarPass` («Cambiar Pass») | botón | `OnClick` | Abre `FormUsuarioCambioPass` | `:215-239`, `.form.txt:1397-1399` |

## Flujos

### Apertura (`Form_Open`, `:336-367`)

1. Si no hay entorno, `EVE` lo inicializa (`:345-347`; `Variables Globales.bas:571`).
2. `EstablecerCombo` rellena `Responsable` con `Constructor.getUsuarios` (`:378-387`).
3. **Sin usuario activo, modo alta:** crea un `usuario` en la global con `ParaAlta = Sí`, administrador «No», activado, no imborrable, caducidad normal, mantener lanzadera abierta, sin permiso de pruebas y `FechaAlta = Date` (`:101-110`).
4. Copia al formulario los campos no vacíos (`:115-187`). Solo si `ParaAlta = No` guarda una copia de partida `m_ObjUsuarioAlInicio` (`:177-183`).
5. Si el usuario **conectado** es administrador, habilita `ComandoRegistrar` y `AllowEdits`; si no, los deshabilita (`:190-196`).
6. `ComandoCambiarPass` solo se habilita si el usuario activo es el conectado (`:197-205`).
7. Cualquier error: `MsgBox` y cierre del formulario (`:358-366`).

### Registrar (`ComandoRegistrar_Click`, `:241-333`)

1. Con copia de partida, `HaHabidoCambios` (`:6-86`); sin cambios, «Ningún cambio que registrar» (`:248-256`).
2. Vuelca los controles al usuario (`:260-307`). `ComboActivado` distinto de «No» activa y vacía `FechaBloqueo`; «No» desactiva y pone `FechaBloqueo = Now()` (`:301-307`).
3. `Usuario.Registrar` (`classes/Usuario.cls:1572-1784`; ver Reglas y Lecturas) (`:308`).
4. Relee el usuario como nueva copia de partida y escribe su `ID` (`:314-318`).

### Errores

`CorreoAlAdministrador` salvo error 1000 y `MsgBox` crítico (`:325-332`). Los mensajes de validación de `Registrar` llegan por `m_Error` (`:309-311`).

## Reglas, validaciones y cálculos

- **Obligatorios:** nombre (`classes/Usuario.cls:1588-1591`), correo (`classes/Usuario.cls:1593-1596`), usuario de red (`classes/Usuario.cls:1597-1600`), matrícula (`classes/Usuario.cls:1601-1604`), `FechaAlta` válida (`classes/Usuario.cls:1605-1608`) y correo con formato válido (`classes/Usuario.cls:1609-1616`).
- **Unicidad** de nombre, correo, usuario de red y matrícula, excluyendo el propio ID si existe (`classes/Usuario.cls:1631-1667, 1546-1550`); mensajes «Ya existe en la base de datos …» (`classes/Usuario.cls:1636, 1646, 1656, 1665`).
- **Alta:** sin ID se genera una contraseña inicial aleatoria y se guarda su resumen criptográfico (`classes/Usuario.cls:1617-1630`); ID con `DameID` (`classes/Usuario.cls:282`); correo en minúsculas (`classes/Usuario.cls:1686`); `TieneQueCambiarLaContrasenia = True` y próximo cambio a 3 meses (`classes/Usuario.cls:1699-1701`); `Observaciones` con el correo del administrador que la registra (`classes/Usuario.cls:1709-1713`).
- **Edición:** la fila se localiza por `CorreoUsuario` (`classes/Usuario.cls:1671-1675`); `Matricula` vacía pasa a `Null` (`classes/Usuario.cls:1729`); `Telefono` y `Movil` vacíos no se escriben (`classes/Usuario.cls:1731-1736`); sin `FechaBloqueo` se fuerza `Activado = True` (`classes/Usuario.cls:1742-1745`).
- **Permiso de edición:** solo administradores conectados (`:190-196`); `AllowEdits` fija la edición manual.
- **Caducidad:** texto según días restantes, «Su Pass nunca caduca» o vacío si el usuario está desactivado (`classes/Usuario.cls:841-879`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo de responsables | `TbUsuariosAplicaciones` (completa) | DAO | no | `modules/Constructor.bas:884-893, 842` |
| Relectura del usuario | `TbUsuariosAplicaciones` por `ID` | SQL dinámico | no | `modules/Constructor.bas:46-73` |
| Unicidad | `TbUsuariosAplicaciones` | SQL dinámico concatenado | no | `classes/Usuario.cls:1546-1554` |
| Alta / edición | `TbUsuariosAplicaciones` | DAO `AddNew`/`Edit` + `Update` | no | `classes/Usuario.cls:1677-1757` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID calculado en VBA con `DameID` (en PostgreSQL, secuencia o identidad); indicadores booleanos mezclados con textos «Sí»/«No» (`EsAdministrador`, `PermisoPruebas`); SQL concatenado sin escapar comillas; sin transacción entre la validación de unicidad y la escritura.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Registrar` no escribe `JefeDelUsuario`, `PermisoPruebas` ni `FechaBaja` en ninguna rama (`classes/Usuario.cls:1677-1757`): el responsable, el permiso de pruebas y la fecha de baja del formulario no se guardarían.
- `ParaAlta` nunca vale `No` (solo se asigna `Sí`, `:103`), así que `m_ObjUsuarioAlInicio` no se carga al abrir (`:177`) y el aviso «Ningún cambio que registrar» no actúa en el primer guardado (`:248`).
- `EsEdicion` solo se asigna a `No` (`classes/Usuario.cls:1721`): el histórico de contraseña y `CorreoAlta`, el correo con la clave inicial, no se ejecutan nunca (`classes/Usuario.cls:1762-1775`). En un alta el usuario no recibiría su clave.
- La edición sobrescribe `FechaAlta` con la fecha actual (`classes/Usuario.cls:1738`).
- Cambiar el correo de un usuario existente no lo localiza (`classes/Usuario.cls:1671-1676`) y entra por `AddNew`, creando otra fila con ID nuevo.
- `HaHabidoCambios` no compara `Telefono`, `Movil` ni `Responsable` (`:25-79`); compara `FechaAlta` y `FechaBaja` como texto con el control (`:67-74`).
- Guardar un usuario desactivado renueva `FechaBloqueo` en cada registro (`:305-306`).
- `Telefono` y `Movil` no se pueden vaciar (`classes/Usuario.cls:1731-1736`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta como acción explícita; clave por identificador estable (ID) y no por correo; enviar la clave inicial por enlace de activación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin nombre, correo, usuario de red, matrícula o fecha de alta, el guardado se rechaza con su mensaje.
2. Un correo, nombre, usuario de red o matrícula repetido en otro usuario se rechaza.
3. Un usuario conectado no administrador ve la ficha sin poder registrar.
4. «Cambiar Pass» solo está habilitado sobre la propia ficha.
5. Desactivar un usuario registra la fecha de bloqueo; reactivarlo la vacía.

## Pendientes

- Contenido de `FormatoCorreoCorrecto`, `DameID` (`Funciones Genreales.bas:455`) y `FechaProximaPassCalculado`.
- Qué campos lee `getUsuario` vía `ColCampos` (`modules/Constructor.bas:80-86`).
- Si el responsable, el permiso de pruebas y la fecha de baja se escriben fuera de `src` (consultas o macros del binario): en `src` no aparece ninguna asignación `Fields("JefeDelUsuario")`, `Fields("PermisoPruebas")` ni `Fields("FechaBaja")`.
