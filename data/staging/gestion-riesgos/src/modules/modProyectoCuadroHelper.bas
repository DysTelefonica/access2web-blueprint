Attribute VB_Name = "modProyectoCuadroHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modProyectoCuadroHelper.bas
'
' Helper para formatear las celdas del cuadro inicial (vista de listado) de
' proyectos en Form_FormProyectosGestion.ListaFiltrados. Punto 17 del acta
' reunion Calidad 2026-06-25.
'
' Helpers publicos (2):
'   ObtenerCeldaFechaUltimaEdicion - devuelve la fecha de publicacion de la
'                                    ultima edicion publicada, o "" si no
'                                    hay. Regla semantica de Natalia:
'                                    "Fecha Edicion = fecha de publicacion".
'                                    Cuando no hay, no se muestra nada.
'   ObtenerCeldaNumeroEdicion     - devuelve el numero de la ultima edicion
'                                    publicada, o "" si no hay.
'
' Decision 2026-07-08 (sesion Punto 17):
'   * Helpers puros: solo reciben datos primitivos o referencias a objetos
'     del modelo (Proyecto, Edicion). No leen DAO directamente. El form es
'     responsable de obtener el proyecto del Dictionary y pasarlo.
'   * Patrón "no mostrar nada" cuando el valor esta vacio: consistente con
'     la regla semantica "Natalia espera que cuando se ponga Fecha de
'     Edicion en cualquier sitio que sea la fecha de publicacion. Cuando no
'     tenga fecha de publicacion alli no pongas nada".
'   * p_Error ByRef por convencion del proyecto (Telefonica D&S).
'
' Reglas de negocio (acordadas con Calidad 2026-07-08 — Punto 17):
'   - Celda "F.Ultima Ed.":
'       * si FechaPublicacion poblada: muestra el valor (dd/mm/yyyy)
'       * si FechaPublicacion vacia: celda vacia (sin "-----", sin label)
'   - Celda "Nº Edicion":
'       * si EdicionUltimaPublicada existe: muestra el numero (string)
'       * si EdicionUltimaPublicada is Nothing: celda vacia
' =============================================================================

' -----------------------------------------------------------------------------
' ObtenerCeldaFechaUltimaEdicion
'
' Devuelve la fecha de la ultima edicion publicada para mostrarla en la celda
' "F.Ultima Ed." del cuadro inicial. Si no hay fecha de publicacion, devuelve
' string vacio (la celda se queda vacia, per regla semantica de Natalia).
'
' Parametros:
'   p_FechaPublicacion (String)        - valor de FechaPublicacion de la
'                                        ultima edicion publicada. Vacio si
'                                        no hay edicion publicada.
'   p_Error (out String)              - mensaje legible si retorna "". Vacio
'                                        si retorna un valor.
'
' Retorna:
'   El valor de p_FechaPublicacion si esta poblado. "" si esta vacio.
' -----------------------------------------------------------------------------
Public Function ObtenerCeldaFechaUltimaEdicion( _
                                            ByVal p_FechaPublicacion As String, _
                                            ByRef p_Error As String _
                                            ) As String
    On Error GoTo errores

    p_Error = ""
    ObtenerCeldaFechaUltimaEdicion = ""

    If Len(Trim$(Nz(p_FechaPublicacion, ""))) = 0 Then
        ' No hay edicion publicada: celda vacia per regla semantica.
        Exit Function
    End If

    ' Hay fecha de publicacion: devolver tal cual (ya viene formateada del modelo).
    ObtenerCeldaFechaUltimaEdicion = Trim$(p_FechaPublicacion)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerCeldaFechaUltimaEdicion: " & Err.Number & " - " & Err.Description
    End If
    ObtenerCeldaFechaUltimaEdicion = ""
End Function

' -----------------------------------------------------------------------------
' ObtenerCeldaNumeroEdicion
'
' Devuelve el numero de la ultima edicion publicada para mostrarlo en la
' celda "Nº Edicion" del cuadro inicial. Si no hay edicion publicada,
' devuelve string vacio (la celda se queda vacia, consistente con la regla
' de "no mostrar nada cuando no hay").
'
' Parametros:
'   p_Edicion (Edicion)               - referencia a la ultima edicion
'                                        publicada. Nothing si no hay.
'   p_Error (out String)              - mensaje legible si retorna "". Vacio
'                                        si retorna un valor.
'
' Retorna:
'   El numero de la edicion (string, viene de Edicion.Edicion). "" si
'   p_Edicion es Nothing.
' -----------------------------------------------------------------------------
Public Function ObtenerCeldaNumeroEdicion( _
                                        ByVal p_Edicion As Edicion, _
                                        ByRef p_Error As String _
                                        ) As String
    On Error GoTo errores

    p_Error = ""
    ObtenerCeldaNumeroEdicion = ""

    If p_Edicion Is Nothing Then
        ' No hay edicion publicada: celda vacia per regla semantica.
        Exit Function
    End If

    ' Hay edicion: devolver el numero (propiedad Edicion del modelo).
    ObtenerCeldaNumeroEdicion = Trim$(Nz(p_Edicion.Edicion, ""))
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerCeldaNumeroEdicion: " & Err.Number & " - " & Err.Description
    End If
    ObtenerCeldaNumeroEdicion = ""
End Function
