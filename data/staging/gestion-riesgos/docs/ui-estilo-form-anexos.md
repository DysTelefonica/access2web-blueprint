# Informe de Estilo UI — Form_FormAnexos

**Proyecto:** GESTION_RIESGOS (00-gestion-riesgos-develop)  
**Objetivo:** Documentar el estilo visual y de controls de Form_FormAnexos y Form_FormAnexos1 para crear dos formularios derivados: (1) Form_FormRiesgosGestionEdicion_Anexos como sub-form en vista Edición, (2) Form_FormRiesgosGestionRiesgo_Anexos como pestaña en vista Riesgo.  
**Solo lectura. No se ha modificado ningún archivo.**  
**Fuente:** Form_FormAnexos.form.txt + Form_FormAnexos.cls + Form_FormAnexos1.form.txt + Form_FormAnexos1.cls + Form_FormRiesgosGestionRiesgo.form.txt + Form_FormRiesgosGestionEdicion.form.txt  
**Inspecciones realizadas con:** dysflow_inspect_form + codegraph-vba (capa código)

---

## 1. Resumen Ejecutivo

Form_FormAnexos y Form_FormAnexos1 son formularios independientes (modal, abierto por `DoCmd.OpenForm` con `openArgs`) que gestionan documentos anexos a tres contextos distintos: **Proyecto**, **Edición** o **Riesgo**. La decisión de contexto se toma en `Form_Open` leyendo `Me.OpenArgs()`. Los eventos personalizados `AnexoAñadido` y `AnexoEliminado` permiten al formulario padre responder a cambios sin acoplamiento estructural.

Para crear los dos formularios derivados:

- **Form_FormRiesgosGestionEdicion_Anexos**: necesita `m_ParaEdicion = EnumSiNo.Sí` + `SourceObject` en un sub-form de `Form_FormRiesgosGestionEdicion` en la zona de detalle ya existente (panel izquierdo con botones `ComandoPublicacion`, `ComandoEliminarEdicion`, etc.).
- **Form_FormRiesgosGestionRiesgo_Anexos**: necesita `m_ParaRiesgo = EnumSiNo.Sí` + `SourceObject` como pestaña adicional en `Form_FormRiesgosGestionRiesgo` (el `lstEstadosHistoricos` y `lblHistorico` están en el detalle de ese formulario; no hay TabControl visible en el árbol de controls, luego la pestaña Histórico podría no existir como control Page/Section diferenciado o bien ser parte del formulario como tal).

> **Nota sobre la pestaña Histórico en `Form_FormRiesgosGestionRiesgo`:** El árbol de controls muestra `lstEstadosHistoricos` (líneas 85-86 del informe de inspección) y `lblHistorico` (línea 84), ambos en la sección Detalle del formulario. No se detecta un PageControl/TabControl explícito en el árbol de controls. La implementación de la pestaña "Histórico" podría ser simplemente la exposición de estos controls en la zona inferior del formulario, sin un contenedor de pestañas formal. **Se recomienda verificar en tiempo de ejecución o revisando el código .cls** si existe un PageControl activo que no aparezca en el `.form.txt`.

---

## 2. Paleta de Colores

| Rol | Color (decimal) | Color (hex) | RGB | Usage |
|---|---|---|---|---|
| Primario (verde corporativo) | 16737792 | 0xFF7E3E | RGB(104, 86, 62) | Header/Footer, labels de texto |
| Fondo detalle | 16774386 | 0xFFE872 | RGB(104, 100, 62) | Fondo de la sección Detalle |
| Fondo alternado lista | 15921906 | 0xF2F2F2 | RGB(242, 242, 242) | Filas alternadas en ListaDocumentos |
| Fondo control de texto editable | 16511723 | 0xFBDD9B | RGB(251, 221, 155) | TextBox editables con `BackColor` |
| Fondo TextBox bloqueado / info | 16446961 | 0xFBBBD1 | RGB(251, 187, 209) | TextBox `Locked` o de solo lectura |
| Fondo ListBox | 16047562 | 0xF4AC96 | RGB(244, 172, 150) | Lista de documentos |
| Borde de control | 14136213 | 0xD75555 | RGB(215, 85, 85) | Bordes de campos editables |
| Texto principal (label) | 16737792 | 0xFF7E3E | RGB(104, 86, 62) | ForeColor de labels `lblTitulo` |
| Texto blanco (label título) | 16777215 | 0xFFFFFF | RGB(255, 255, 255) | `lblTitulo1` en header |
| Texto de control editable | 8210719 | 0x7D5E3F | RGB(125, 94, 63) | ForeColor de TextBox y ListBox |
| Botón eliminar | 2366701 | 0x242425 | RGB(36, 36, 37) | `ComandoEliminarAnexo` (texto rojo oscuro) |
| Hover botón comando | 15060409 | 0xE5E5E5 | RGB(229, 229, 229) | `HoverColor` en CommandButton |
| Pressed botón comando | 9592887 | 0x925C5C | RGB(146, 92, 92) | `PressedColor` en CommandButton |

---

## 3. Tipografía

| Control | Font | Size | Weight | Usage |
|---|---|---|---|---|
| Título del formulario (`lblTitulo1`) | Segoe UI | 20 | 700 (Bold) | Título "ANEXOS PARA..." en header |
| Etiquetas de control (`lblTitulo`, `lblEvidenciaUTE`) | Segoe UI | 14 | 700 (Bold) | Labels junto a campos editables |
| Campos editables (`TituloDoc`, `txtRuta`) | Segoe UI | 14 | 400 (Normal) | TextBox con `BackColor = 16511723` |
| Campo observaciones (`Titulo` multilínea) | Segoe UI | 14 | 400 (Normal) | TextBox con `ScrollBars = 2` |
| Lista de documentos (`ListaDocumentos`) | Segoe UI | — | — | ListBox, Font heredado |
| Botones comando | Segoe UI | 11 | 400 (Normal) | CommandButton estándar |
| Botón eliminar (`ComandoEliminarAnexo`) | Segoe UI | 10 | 700 (Bold) | Bold, color `ForeColor = 2366701` |
| Botones de encabezado/pie | — | — | 700 (Bold) | ComandoAyuda, cmdSalir |

---

## 4. Dimensiones Principales (twips)

### Form_FormAnexos y Form_FormAnexos1 (idénticas)

| Dimensión | Valor (twips) | Equivalente aprox. (cm) |
|---|---|---|
| Ancho total (`Width`) | 13.722 | 9,6 cm |
| FormHeader altura (`Height`) | 686 | 0,48 cm |
| Detalle altura (`Height`) | 6.803 | 4,75 cm |
| FormFooter altura (`Height`) | 686 | 0,48 cm |
| Alto total aprox. | 8.175 + márgenes | ≈ 5,71 cm |

### Form_FormRiesgosGestionRiesgo

| Dimensión | Valor (twips) | Uso |
|---|---|---|
| Ancho total (`Width`) | 11.917 | Formulario principal |
| Detalle altura (`Height`) | 6.633 | Zona principal del formulario |

### Form_FormRiesgosGestionEdicion

| Dimensión | Valor (twips) | Uso |
|---|---|---|
| Ancho total (`Width`) | 11.926 | Formulario principal |
| Detalle altura (`Height`) | 6.576 | Zona principal del formulario |

---

## 5. Estructura de Controles

### 5.1 Form_FormAnexos1 (formulario origen de referencia, sin txtRuta/btnAceptar)

**Encabezado (FormHeader: alto=686, verde corporativo)**

```
lblTitulo1       (Label)   left=0, top=0,     width=11820, height=480  — Título "ANEXOS PARA EL RIESGO: R001"
ImagenAnexo      (Image)    left=11985, top=30, width=618,  height=618  — Icono PDF, OnClick abre archivo
ComandoAyuda     (Button)  left=12768, top=30, width=618,  height=618  — "Salir"
```

**Detalle (Detail: alto=6803)**

```
ListaDocumentos  (ListBox)  left=345,  top=2085, width=13041, height=3750 — Lista de anexos
  ColumnCount=6, ColumnWidths="737;0;7718;1134;1701;1451"
  RowSource="Tipo;IDAnexo;Título;Riesgo;Edic_Anexado;Fecha"

lblEvidenciaUTE   (Label)   left=350,  top=255,  width=5010, height=390  — "Es evidencia tratamiento riesgos con UTE"
EvidenciaUTE      (ComboBox) left=5510, top=255,  width=1236, height=390  — "Sí;No" (visible si EnUTE="Sí")

lblTitulo        (Label)   left=350,  top=885,  width=915,  height=390  — "Título"
TituloDoc        (TextBox)  left=1418, top=885,  width=11447, height=390 — Campo de título del anexo

ComandoAnexar    (Button)  left=12995, top=1425, width=391, height=390  — "Comando47" (icono carpeta)

Titulo           (TextBox) left=350,  top=5985, width=12006, height=672 — Observaciones (bloqueado, multilínea)
ComandoEliminarAnexo (Button) left=12456, top=6236, width=930, height=390 — "Borrar" (rojo)
```

**Pie (FormFooter: alto=686, verde corporativo)**

```
cmdSalir (Button) left=12808, top=45, width=578, height=578 — "Salir"
```

### 5.2 Form_FormAnexos (formulario con más controles UI, diff vs. Form_FormAnexos1)

**Diferencias respecto a Form_FormAnexos1 en el detalle:**

| Control extra en Form_FormAnexos | Descripción | Líneas .cls |
|---|---|---|
| `txtRuta` (TextBox) | Muestra la ruta del archivo seleccionado tras pulsar btnExaminar | 48 |
| `btnExaminar` (CommandButton) | Abre el diálogo de selección de archivo (`msoFileDialogFilePicker`) | 52–91 |
| `btnAceptar` (CommandButton) | Valida, persiste el anexo con `Anexo.Registrar`, lanza `AnexoAñadido` | 93–187 |
| `btnActualizarNombreArchivo` (CommandButton) | Renombra título del anexo seleccionado en `m_ObjAnexoSeleccionado.CambiarNombre` | 189–227 |
| `Etiqueta7` (Label) | Label para `txtRuta` ("ruta") | 47 |

**Eventos en Form_FormAnexos.cls (líneas 93–227):**
- `btnExaminar_Click()` → `Application.FileDialog(3)` + Sugiere nombre de archivo base
- `btnAceptar_Click()` → Validaciones + duplicados (`ExisteNombreEnAmbito`) + `RaiseEvent AnexoAñadido`
- `btnActualizarNombreArchivo_Click()` → `m_ObjAnexoSeleccionado.CambiarNombre`
- `ComandoEliminarAnexo_Click()` → `RaiseEvent AnexoEliminado`
- `ListaDocumentos_Click()` → Muestra `txtRuta`, habilita `btnActualizarNombreArchivo`

---

## 6. Lista de Anexos — Especificación Técnica

### Columnas (ListaDocumentos)

| Índice | Ancho (twips) | Ancho real mostrado | Contenido |
|---|---|---|---|
| 0 | 737 | Visible (Tipo) | "R" (riesgo) o "E" (edición) |
| 1 | 0 | Oculta | IDAnexo (clave primaria) |
| 2 | 7.718 | Visible (ancho principal) | Título del anexo |
| 3 | 1.134 | Visible | Código de riesgo (`m_ObjAnexo.riesgo.CodigoRiesgo`) o vacío |
| 4 | 1.701 | Visible | Edición en la que se anexó (`Edicion.Edicion`) |
| 5 | 1.451 | Visible | Fecha del anexo (`Format(.FechaAnexo, "dd/mm/yyyy")`) |

**Ancho total visible de columnas:** 737 + 0 + 7718 + 1134 + 1701 + 1451 = **13.241 twips** (dentro del ListBox de ancho total 13.041 + left 345 = 13.386 — margen de ~145 twips).

**Delegación de datos** (`Form_FormAnexos1.cls:71-72`):

```vba
If m_ParaProyecto = EnumSiNo.Sí Then
    Set m_ObjColAnexos = m_ObjDELAnexo.ColAnexosTotales
Else
    Set m_ObjColAnexos = m_ObjDELAnexo.ColAnexos
End If
```

### Eventos de Selection

| Evento | Control | Action |
|---|---|---|
| `OnClick` | ListaDocumentos | Selecciona anexo, muestra `Titulo` (observaciones), habilita `ComandoEliminarAnexo` si `blnPermitidoEscribir = True` |
| `OnDblClick` | ListaDocumentos | Si `ImagenAnexo.Visible = True` → `ImagenAnexo_Click` → `AbrirEnLocal m_URLAnexoSeleccionado` |

---

## 7. Eventos Personalizados (Custom Events) — Punto de Integración con Padres

Ambos formularios (`Form_FormAnexos1` y `Form_FormAnexos`) declaran los mismos dos eventos personalizados (líneas 15–16 de Form_FormAnexos1.cls, líneas 22–24 de Form_FormAnexos.cls):

```vba
Public Event AnexoAñadido(ByVal p_ObjAnexo As Anexo)
Public Event AnexoEliminado(ByVal IDAnexo As String)
```

### Firma de los eventos

| Evento | Parámetro | Tipo | Contenido |
|---|---|---|---|
| `AnexoAñadido` | `p_ObjAnexo` | `Anexo` | Objeto recién creado y persistido |
| `AnexoEliminado` | `IDAnexo` | `String` | ID del anexo eliminado |

### Integración en el padre (patrón de uso en el código existente)

```vba
' En el formulario padre (ej. Form_FormRiesgosGestionRiesgo):
Private WithEvents m_frmAnexos As Form_FormAnexos1

Private Sub m_frmAnexos_AnexoAñadido(ByVal p_ObjAnexo As Anexo)
    ' Refrescar lista de anexos o actualizar contador
    Me.EstablecerLista
End Sub

Private Sub m_frmAnexos_AnexoEliminado(ByVal IDAnexo As String)
    ' Refrescar lista de anexos
    Me.EstablecerLista
End Sub
```

### Cómo se invocan los eventos

| Evento | Se lanza desde | Líneas .cls |
|---|---|---|
| `AnexoAñadido` | `ComandoAnexar_Click()` / `btnAceptar_Click()` tras `.Registrar` exitoso | 217 (Form_FormAnexos1.cls), 168 (Form_FormAnexos.cls) |
| `AnexoEliminado` | `ComandoEliminarAnexo_Click()` tras `.EliminarAnexo` exitoso | 306 (Form_FormAnexos1.cls), 378 (Form_FormAnexos.cls) |

---

## 8. Constantes de Diseño (Alturas en twips)

Definidas en `Form_FormAnexos1.cls` y replicadas en `Form_FormAnexos.cls`:

| Constante | Valor twips | Contexto |
|---|---|---|
| `ALTO_FORM_CONUTE` | 9.145 | Alto total con evidencia UTE visible |
| `ALTO_FORM_SINUTE` | 7.836 | Alto total sin evidencia UTE |
| `SUPERIOR_ANEXO_CONUTE` | 1.679 | Posición top del campo Anexo/Urgente (con UTE) |
| `SUPERIOR_ANEXO_SINUTE` | 1.049 | Posición top del campo Anexo/Urgente (sin UTE) |
| `ALTO_CUADRO_PPAL_CONUTE` | 7.753 | Alto del cuadro principal (con UTE) |
| `ALTO_CUADRO_PPAL_SINUTE` | 6.464 | Alto del cuadro principal (sin UTE) |
| `SUPERIOR_LISTA_CONUTE` | 2.384 | Posición top de ListaDocumentos (con UTE) |
| `SUPERIOR_LISTA_SINUTE` | 959 | Posición top de ListaDocumentos (sin UTE) |
| `SUPERIOR_TITULO_CONUTE` | 7.709 | Posición top del campo observaciones (con UTE) |
| `SUPERIOR_TITULO_SINUTE` | 6.284 | Posición top del campo observaciones (sin UTE) |
| `SUPERIOR_CUADRO2_CONUTE` | 8.459 | Posición top del botón "Salir" (con UTE) |
| `SUPERIOR_CUADRO2_SINUTE` | 7.150 | Posición top del botón "Salir" (sin UTE) |
| `SUPERIOR_SALIR_CONUTE` | 8.518 | Posición top botón salir (con UTE) |
| `SUPERIOR_SALIR_SINUTE` | 7.228 | Posición top botón salir (sin UTE) |

### Lógica de adaptación de tamaño (`AdaptarTamañoFormulario`)

El formulario ajusta la visibilidad y posición de los controles según el flag `m_EnUTE`:

```vba
If m_EnUTE = "Sí" Then
    Me.lblEvidenciaUTE.Visible = True
    Me.EvidenciaUTE.Visible = True
    ' → usa constantes CONUTE
Else
    Me.lblEvidenciaUTE.Visible = False
    Me.EvidenciaUTE.Visible = False
    ' → usa constantes SINUTE
End If
```

---

## Anexo: Comparación de Controles entre Form_FormAnexos1 y Form_FormAnexos

| Control | Form_FormAnexos1 | Form_FormAnexos | Notas |
|---|---|---|---|
| FormHeader | ✓ | ✓ | Verde corporativo, alto 686 |
| `lblTitulo1` | ✓ | ✓ | Título dinámico según openArgs |
| `ImagenAnexo` | ✓ (Visible=NotDefault) | ✓ (visible por defecto) | OnClick abre archivo |
| `ComandoAyuda` | ✓ | ✓ | "Salir" |
| `Detalle` | alto=6803 | alto=6803 | |
| `EvidenciaUTE` ComboBox | Visible/Enabled según UTE | Igual | |
| `lblEvidenciaUTE` | Visible=NotDefault | Visible=NotDefault | |
| `TituloDoc` | width=11447 | width=9407 (más estrecho) | Form_FormAnexos tiene `txtRuta` al lado |
| `lblTitulo` | ✓ | ✓ | |
| `ComandoAnexar` / `btnExaminar` | Nombre interno "Comando47" | Nombre interno "Comando47" | Mismo icono CarpetaAbiertaRoja.ico |
| `txtRuta` | ✗ | ✓ | TextBox con la ruta seleccionada |
| `btnAceptar` | ✗ | ✓ | Botón "Aceptar" |
| `btnActualizarNombreArchivo` | ✗ | ✓ | Visible=NotDefault, solo si hay selección |
| `ListaDocumentos` | ColumnCount=6 | ColumnCount=6 | Misma estructura |
| `Titulo` | multilínea bloqueado | multilínea bloqueado | |
| `ComandoEliminarAnexo` | ForeColor rojo 2366701 | ForeColor rojo 2366701 | |
| FormFooter | alto=686, verde | alto=686, verde | |
| `cmdSalir` | ✓ | ✓ | |
| Eventos personalizados | `AnexoAñadido`, `AnexoEliminado` | `AnexoAñadido`, `AnexoEliminado` | Idénticos |
