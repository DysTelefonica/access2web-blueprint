Attribute VB_Name = "SandboxCloneHelper"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: SandboxCloneHelper
' DESCRIPCIÓN: Motor DAO de bajo nivel para clonado de tablas linked ? local
'              en el sandbox de testing. Operaciones puras de datos,
'              sin lógica de lifecycle.
' RESPONSABILIDAD:
'   - Detectar si una tabla es linked
'   - Obtener su TableDef
'   - Crear una tabla local compatible
'   - Copiar datos desde linked a local
'   - Eliminar/desvincular la linked original del sandbox
' FASE: 2 del cambio sandbox-linked-localization
' ==========================================================================

' --- Constantes de errores ---
Private Const ERR_NOT_LINKED As Long = 20010
Private Const ERR_TABLE_NOT_FOUND As Long = 20011
Private Const ERR_UNSUPPORTED_FIELD As Long = 20012
Private Const ERR_COPY_FAILED As Long = 20013
Private Const ERR_DELETE_FAILED As Long = 20014
Private Const ERR_LOCAL_TABLE_NOT_FOUND As Long = 20015
Private Const ERR_REPLACE_FAILED As Long = 20016

' --- Constantes DAO para detección de linked tables ---
Private Const DB_LINKED As Long = &H80000000  ' dbAttached flag

' --- Tamaño de batch para transacciones ---
Private Const BATCH_SIZE As Long = 100

' ==========================================================================
' MÉTODOS PÚBLICOS
' ==========================================================================

' ---
' IsLinkedTable: Detecta si una tabla es linked en la base de datos dada
' tableName: Nombre de la tabla a verificar
' dbTarget: Database donde buscar (si Nothing, usa CurrentDb)
' Retorna: True si la tabla existe y es linked
' Lanza: CondorError si la tabla no existe
' ---
Public Function IsLinkedTable(tableName As String, Optional dbTarget As DAO.Database = Nothing) As Boolean
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Buscar la tabla
    Set tdf = db.TableDefs(tableName)
    
    ' Verificar si es tabla de sistema (omitir)
    If Left(tdf.name, 4) = "MSys" Or Left(tdf.name, 4) = "USys" Then
        IsLinkedTable = False
        Exit Function
    End If
    
    ' Detectar si es linked: el flag dbAttached está presente
    ' En DAO, las tablas linked tienen el atributo dbAttached (0x80000000)
    IsLinkedTable = IsLinkedTableDef(tdf)
    
    Exit Function
    
Errores:
    If Err.Number = 3265 Then  ' Table not found
        ' FIX: En lugar de lanzar error, retornar False.
        ' Esto permite que el caller distinga "no existe" de "existe y es linked".
        ' El caller (CloneTableToSandbox) usa Resume Next, y si lanzamos error
        ' aqui, el comportamiento es impredecible - a veces el error se propaga,
        ' a veces no (Access/VBA quirk). Retornar False es el comportamiento correcto.
        IsLinkedTable = False
        Exit Function
    End If
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxCloneHelper.IsLinkedTable"
    errObj2.Raise
End Function

' ---
' TableExists: Verifica si una tabla existe en la base de datos dada
' tableName: Nombre de la tabla a verificar
' dbTarget: Database donde buscar (si Nothing, usa CurrentDb)
' Retorna: True si la tabla existe (sea linked o local)
' ---
Public Function TableExists(tableName As String, Optional dbTarget As DAO.Database = Nothing) As Boolean
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Buscar la tabla
    Set tdf = db.TableDefs(tableName)
    TableExists = True
    Exit Function
    
Errores:
    TableExists = False
End Function

' ---
' GetSourceTableDef: Obtiene el TableDef de una tabla en el backend,
' sea LINKED o LOCAL. Usada por CreateLocalTableFromLinked para resolver
' la estructura del origen sin importar su tipo.
' tableName: Nombre de la tabla
' dbTarget: Database donde buscar (si Nothing, usa CurrentDb)
' Retorna: DAO.TableDef de la tabla
' Lanza: CondorError si la tabla no existe
' ---
Public Function GetSourceTableDef(tableName As String, Optional dbTarget As DAO.Database = Nothing) As DAO.TableDef
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Obtener TableDef (acepta linked O local)
    Set tdf = db.TableDefs(tableName)
    
    Set GetSourceTableDef = tdf
    Exit Function
    
Errores:
    If Err.Number = 3265 Then  ' Table not found
        Dim errObj As New CondorError
        errObj.Create ERR_TABLE_NOT_FOUND, "Tabla no encontrada: " & tableName, "SandboxCloneHelper.GetSourceTableDef"
        errObj.Raise
    End If
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxCloneHelper.GetSourceTableDef"
    errObj2.Raise
End Function

' ---
' GetLinkedTableDef: Obtiene el TableDef de una tabla linked
' tableName: Nombre de la tabla linked
' dbTarget: Database donde buscar (si Nothing, usa CurrentDb)
' Retorna: DAO.TableDef de la tabla linked
' Lanza: CondorError si no existe o no es linked
' ---
Public Function GetLinkedTableDef(tableName As String, Optional dbTarget As DAO.Database = Nothing) As DAO.TableDef
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Obtener TableDef
    Set tdf = db.TableDefs(tableName)
    
    ' Verificar que es linked
    If Not IsLinkedTableDef(tdf) Then
        Dim errObj As New CondorError
        errObj.Create ERR_NOT_LINKED, "La tabla '" & tableName & "' no es una tabla linked", "SandboxCloneHelper.GetLinkedTableDef"
        errObj.Raise
    End If
    
    Set GetLinkedTableDef = tdf
    Exit Function
    
Errores:
    If Err.Number = 3265 Then  ' Table not found
        Dim errObj2 As New CondorError
        errObj2.Create ERR_TABLE_NOT_FOUND, "Tabla linked no encontrada: " & tableName, "SandboxCloneHelper.GetLinkedTableDef"
        errObj2.Raise
    End If
    Dim errObj3 As New CondorError
    errObj3.Create Err.Number, Err.description, "SandboxCloneHelper.GetLinkedTableDef"
    errObj3.Raise
End Function

' ---
' CreateLocalTableFromLinked: Crea una tabla LOCAL en el sandbox con la estructura
'                             de la tabla linked de origen (SIN COPIAR DATOS)
' sourceTableName: Nombre de la tabla linked de origen (para obtener estructura)
' tempLocalTableName: Nombre para la nueva tabla local
' dbSource: Database del backend donde está la tabla linked (requerido para cross-db)
' dbTarget: Database del sandbox donde crear la tabla (si Nothing, usa CurrentDb)
' Lanza: CondorError si la tabla origen no es linked o tiene campos no soportados
' ---
Public Sub CreateLocalTableFromLinked(sourceTableName As String, tempLocalTableName As String, dbSource As DAO.Database, Optional dbTarget As DAO.Database = Nothing)
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdfSource As DAO.TableDef
    Dim tdfNew As DAO.TableDef
    Dim fldSource As DAO.Field
    Dim idxSource As DAO.Index
    Dim fldNew As DAO.Field
    Dim idxNew As DAO.Index
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Obtener TableDef de la tabla origen (linked o local) desde el backend
    Set tdfSource = GetSourceTableDef(sourceTableName, dbSource)
    
    ' Crear nueva TableDef
    Set tdfNew = db.CreateTableDef(tempLocalTableName)
    
    ' Copiar campos
    For Each fldSource In tdfSource.Fields
        ' Verificar tipo de campo
        Select Case fldSource.Type
            Case dbLongBinary, dbAttachment, dbByte, dbVarBinary, dbGUID
                ' Verificar si es un campo OLE/Attachment/MultiValue no soportado
                If Not IsSupportedFieldType(fldSource.Type) Then
                    Dim errObj As New CondorError
                    errObj.Create ERR_UNSUPPORTED_FIELD, _
                        "Campo no soportado en clonación: '" & fldSource.name & "' (tipo: " & fldSource.Type & ") en tabla '" & sourceTableName & "'. Solo se soportan tipos de datos estándar.", _
                        "SandboxCloneHelper.CreateLocalTableFromLinked"
                    errObj.Raise
                End If
        End Select
        
        ' Crear campo en la nueva tabla
        Set fldNew = CreateFieldFromSource(fldSource, tdfNew)
        
    Next fldSource
    
    ' Copiar índices (incluyendo primary key)
    For Each idxSource In tdfSource.Indexes
        Set idxNew = tdfNew.CreateIndex(idxSource.name)
        idxNew.Primary = idxSource.Primary
        idxNew.Unique = idxSource.Unique
        idxNew.Required = idxSource.Required
        idxNew.IgnoreNulls = idxSource.IgnoreNulls
        
        ' Copiar campos del índice
        Dim fldIdx As DAO.Field
        For Each fldIdx In idxSource.Fields
            idxNew.Fields.Append idxNew.CreateField(fldIdx.name)
        Next fldIdx
        
        tdfNew.Indexes.Append idxNew
    Next idxSource
    
    ' Agregar la tabla a la database
    db.TableDefs.Append tdfNew
    
    Exit Sub
    
Errores:
    Dim errObjFinal As New CondorError
    errObjFinal.Create Err.Number, Err.description, "SandboxCloneHelper.CreateLocalTableFromLinked"
    errObjFinal.Raise
End Sub

' ---
' CopyTableData: Copia datos de una tabla origen a una tabla destino
'               USA batch transactions de 100 filas para robustez
' sourceTableName: Nombre de la tabla origen (debe existir en dbSource)
' targetTableName: Nombre de la tabla destino (debe existir con estructura compatible en dbTarget)
' dbSource: Database del backend donde está la tabla linked (requerido para cross-db)
' dbTarget: Database del sandbox donde está la tabla local (si Nothing, usa CurrentDb)
' Retorna: Cantidad de filas copiadas
' Lanza: CondorError si falla la copia
' ---
Public Function CopyTableData(sourceTableName As String, targetTableName As String, dbSource As DAO.Database, Optional dbTarget As DAO.Database = Nothing) As Long
    On Error GoTo ErroresTransaccion
    
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rsSrc As DAO.Recordset
    Dim rsDst As DAO.Recordset
    Dim fld As DAO.Field
    Dim lngRowCount As Long
    Dim lngBatchCount As Long
    
    ' Usar database especificada o CurrentDb para el destino
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    Set ws = DBEngine.Workspaces(0)
    
    ' Abrir recordset de origen desde el backend (linked table - Access sigue el Connect string)
    Set rsSrc = dbSource.OpenRecordset("SELECT * FROM [" & sourceTableName & "]", dbOpenForwardOnly)
    
    ' Abrir recordset de destino para edición (tabla local en el sandbox)
    Set rsDst = db.OpenRecordset(targetTableName, dbOpenDynaset)
    
    ' Iniciar transacción
    lngRowCount = 0
    lngBatchCount = 0
    ws.BeginTrans
    
    Do While Not rsSrc.EOF
        rsDst.AddNew
        
        For Each fld In rsSrc.Fields
            ' Omitir campos no soportados (OLE, Attachment, etc)
            If IsCopyableField(fld.Type) Then
                On Error Resume Next
                rsDst.Fields(fld.name).value = fld.value
                If Err.Number <> 0 Then
                    ' Campo puede no existir en destino o tipo incompatible
                    ' Omitir silenciosamente - la estructura debe ser compatible
                    Err.Clear
                End If
                On Error GoTo ErroresTransaccion
            End If
        Next fld
        
        rsDst.Update
        lngRowCount = lngRowCount + 1
        lngBatchCount = lngBatchCount + 1
        
        ' Commit por batch
        If lngBatchCount >= BATCH_SIZE Then
            ws.CommitTrans dbForceOSFlush
            ws.BeginTrans
            lngBatchCount = 0
        End If
        
        rsSrc.MoveNext
    Loop
    
    ' Commit final
    ws.CommitTrans dbForceOSFlush
    
    rsSrc.Close
    rsDst.Close
    Set rsSrc = Nothing
    Set rsDst = Nothing
    
    CopyTableData = lngRowCount
    
    Exit Function
    
ErroresTransaccion:
    On Error GoTo Errores
    ws.Rollback
    
    Dim errObj As New CondorError
    errObj.Create ERR_COPY_FAILED, "Error al copiar datos de '" & sourceTableName & "' a '" & targetTableName & "': " & Err.description & " (filas copiadas antes del error: " & lngRowCount & ")", "SandboxCloneHelper.CopyTableData"
    errObj.Raise
    Exit Function
    
Errores:
    Dim errObj2 As New CondorError
    errObj2.Create Err.Number, Err.description, "SandboxCloneHelper.CopyTableData"
    errObj2.Raise
End Function

' ---
' DeleteLinkedTable: Elimina una tabla linked del sandbox
' tableName: Nombre de la tabla linked a eliminar
' dbTarget: Database del sandbox (si Nothing, usa CurrentDb)
' Lanza: CondorError si no es linked o no existe
' ---
Public Sub DeleteLinkedTable(tableName As String, Optional dbTarget As DAO.Database = Nothing)
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' Verificar que existe y es linked
    If Not IsLinkedTable(tableName, db) Then
        Dim errObj As New CondorError
        errObj.Create ERR_NOT_LINKED, "La tabla '" & tableName & "' no es una tabla linked. No se puede eliminar.", "SandboxCloneHelper.DeleteLinkedTable"
        errObj.Raise
    End If
    
    ' Eliminar la linked table
    ' NOTA: Access preguntará al usuario si desea eliminar el vínculo
    '       En código automatizado, esto puede requerir confirmación
    db.TableDefs.Delete tableName
    
    Exit Sub
    
Errores:
    If Err.Number = 3265 Then  ' Table not found
        Dim errObj2 As New CondorError
        errObj2.Create ERR_TABLE_NOT_FOUND, "Tabla no encontrada: " & tableName, "SandboxCloneHelper.DeleteLinkedTable"
        errObj2.Raise
    End If
    Dim errObj3 As New CondorError
    errObj3.Create Err.Number, Err.description, "SandboxCloneHelper.DeleteLinkedTable"
    errObj3.Raise
End Sub

' ---
' ReplaceLinkedWithLocal: Reemplaza una tabla linked por su versión local
'                         Haciendo el switch seguro para que la tabla final
'                         se llame EXACTAMENTE igual que la original
' tableName: Nombre original de la tabla linked (nombre final deseado)
' tempLocalTableName: Nombre de la tabla local temporal que reemplazará a la linked
' dbTarget: Database del sandbox (si Nothing, usa CurrentDb)
' Lanza: CondorError si algo falla en el proceso
' Flujo:
'   1. Validar que tableName es linked
'   2. Validar que tempLocalTableName existe y es local
'   3. Eliminar linked original
'   4. Renombrar la tabla local temporal al nombre final tableName
' ---
Public Sub ReplaceLinkedWithLocal(tableName As String, tempLocalTableName As String, Optional dbTarget As DAO.Database = Nothing)
    On Error GoTo Errores
    
    Dim db As DAO.Database
    Dim tdfLocal As DAO.TableDef
    Dim blnLocalEsLinked As Boolean
    
    ' Usar database especificada o CurrentDb
    If dbTarget Is Nothing Then
        Set db = CurrentDb
    Else
        Set db = dbTarget
    End If
    
    ' PASO 1: Validar que tableName es linked
    If Not IsLinkedTable(tableName, db) Then
        Dim errObj As New CondorError
        errObj.Create ERR_NOT_LINKED, "La tabla '" & tableName & "' no es una tabla linked. No se puede reemplazar.", "SandboxCloneHelper.ReplaceLinkedWithLocal"
        errObj.Raise
    End If
    
    ' PASO 2: Validar que tempLocalTableName existe y es local
    ' FIX: Refrescar TableDefs antes de buscar la tabla recien creada en el sandbox.
    ' Sin esto, Access puede no ver la tabla temp aunque ya fue creada.
    db.TableDefs.Refresh
    On Error Resume Next
    Set tdfLocal = db.TableDefs(tempLocalTableName)
    If Err.Number <> 0 Then
        On Error GoTo Errores
        Dim errObj2 As New CondorError
        errObj2.Create ERR_LOCAL_TABLE_NOT_FOUND, "Tabla local temporal '" & tempLocalTableName & "' no encontrada en el sandbox", "SandboxCloneHelper.ReplaceLinkedWithLocal"
        errObj2.Raise
    End If
    On Error GoTo Errores
    
    ' Verificar que NO es linked
    blnLocalEsLinked = IsLinkedTableDef(tdfLocal)
    If blnLocalEsLinked Then
        Dim errObj3 As New CondorError
        errObj3.Create ERR_UNSUPPORTED_FIELD, "La tabla '" & tempLocalTableName & "' es linked. Debe ser una tabla local.", "SandboxCloneHelper.ReplaceLinkedWithLocal"
        errObj3.Raise
    End If
    
    ' Verificar que no es la misma tabla (evitar accidental rename sobre sí misma)
    If tableName = tempLocalTableName Then
        ' Ya está reemplazada, no hacer nada
        Exit Sub
    End If
    
    ' PASO 3: Eliminar linked original
    db.TableDefs.Delete tableName
    
    ' PASO 4: Renombrar la tabla local temporal al nombre final
    tdfLocal.name = tableName
    
    Exit Sub
    
Errores:
    Dim errObjFinal As New CondorError
    errObjFinal.Create Err.Number, Err.description, "SandboxCloneHelper.ReplaceLinkedWithLocal"
    errObjFinal.Raise
End Sub

' ==========================================================================
' MÉTODOS PRIVADOS (Helpers internos)
' ==========================================================================

' ---
' IsLinkedTableDef: Evalúa si un TableDef representa una tabla linked
' Usa el flag dbAttached (&H80000000) en Attributes
' ---
Private Function IsLinkedTableDef(tdf As DAO.TableDef) As Boolean
    On Error GoTo Errores
    
    ' Las tablas linked tienen el bit dbAttached set en Attributes
    ' dbAttached = &H80000000 = 2147483648
    IsLinkedTableDef = (tdf.Attributes And DB_LINKED) = DB_LINKED
    
    Exit Function
    
Errores:
    IsLinkedTableDef = False
End Function

' ---
' IsSupportedFieldType: Verifica si el tipo de campo DAO es soportado
' para clonación. OLE Object, Attachment, MultiValue = NO soportados.
' ---
Private Function IsSupportedFieldType(fieldType As Integer) As Boolean
    Select Case fieldType
        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDate, _
             dbText, dbMemo, dbBoolean, dbBigInt, dbVarBinary, dbGUID
            IsSupportedFieldType = True
        Case Else
            IsSupportedFieldType = False
    End Select
End Function

' ---
' IsCopyableField: Verifica si el tipo de campo puede ser copiado
'Diferentes criterios que IsSupportedFieldType para creación de estructura
' ---
Private Function IsCopyableField(fieldType As Integer) As Boolean
    Select Case fieldType
        Case dbLongBinary, dbAttachment
            ' Campos OLE y Attachment se omiten en la copia
            IsCopyableField = False
        Case Else
            IsCopyableField = True
    End Select
End Function

' ---
' CreateFieldFromSource: Crea un DAO.Field en la tabla destino basándose
'                        en el campo origen, manejando atributos especiales
' ---
Private Function CreateFieldFromSource(fldSource As DAO.Field, tdfParent As DAO.TableDef) As DAO.Field
    On Error GoTo Errores
    
    Dim fldNew As DAO.Field
    Dim lngFieldType As Long
    Dim lngSize As Long
    
    lngFieldType = fldSource.Type
    
    ' Determinar tamaño del campo
    ' Para Text, el tamaño está en Attributes (valores altos)
    ' Para otros tipos, usar el tamaño nativo o 0
    Select Case lngFieldType
        Case dbText
            ' Tamaño para texto: está codificado en los attributes altos
            ' O usar el valor de la propiedad Size si está disponible
            lngSize = GetFieldSize(fldSource)
        Case dbMemo
            lngSize = 0  ' Memo no tiene límite en CreateField
        Case Else
            lngSize = 0
    End Select
    
    ' Crear el campo
    If lngSize > 0 Then
        Set fldNew = tdfParent.CreateField(fldSource.name, lngFieldType, lngSize)
    Else
        Set fldNew = tdfParent.CreateField(fldSource.name, lngFieldType)
    End If
    
    ' Copiar atributos de nullable/required
    On Error Resume Next
    fldNew.Required = fldSource.Required
    fldNew.AllowZeroLength = fldSource.AllowZeroLength
    On Error GoTo Errores
    
    ' Copiar propiedades extendidas si existen
    Call CopyFieldProperties(fldSource, fldNew)
    
    ' Agregar a la colección de campos
    tdfParent.Fields.Append fldNew
    
    Set CreateFieldFromSource = fldNew
    Exit Function
    
Errores:
    Set CreateFieldFromSource = Nothing
End Function

' ---
' GetFieldSize: Obtiene el tamaño de un campo de texto desde sus atributos
' Los campos Text en Access almacenan el tamaño en los 16 bits altos de Attributes
' ---
Private Function GetFieldSize(fld As DAO.Field) As Long
    On Error GoTo Errores
    
    Dim lngAttr As Long
    Dim lngSize As Long
    
    lngAttr = fld.Attributes
    
    ' Para campos de texto, el tamaño está codificado:
    ' Los 16 bits altos de Attributes contienen el tamaño (solo para dbText)
    ' Si el bit 1 (dbFixedField) está set, es de longitud fija
    If (lngAttr And dbFixedField) = dbFixedField Then
        ' Longitud fija
        lngSize = lngAttr \ &H10000 And &HFF&
    Else
        ' Longitud variable: los bits superiores contienen el tamaño
        ' En Access, los campos Text variables tienen el tamaño en bits 16-23
        lngSize = lngAttr \ &H10000 And &HFFFF&
        If lngSize = 0 Then lngSize = 255  ' Default Access
    End If
    
    GetFieldSize = lngSize
    Exit Function
    
Errores:
    GetFieldSize = 255  ' Default
End Function

' ---
' CopyFieldProperties: Copia propiedades extendidas del campo origen al destino
' Maneja errores suavemente (propiedades pueden no existir)
' ---
Private Sub CopyFieldProperties(fldSource As DAO.Field, fldDest As DAO.Field)
    On Error Resume Next
    
    Dim prp As DAO.Property
    
    ' Copiar propiedades estándar que podrían existir
    For Each prp In fldSource.Properties
        Select Case prp.name
            Case "OrdinalPosition", "ColumnHidden", "ColumnOrder"
                ' Propiedades de visualización, ignorar
            Case Else
                ' Intentar copiar la propiedad
                On Error Resume Next
                fldDest.Properties(prp.name) = prp.value
                On Error GoTo 0
        End Select
    Next prp
    
    Set prp = Nothing
End Sub

