Option Explicit

' Variables para rastrear cambios y permitir deshacer
Private filasModificadas As New Collection

Private Function ConvertirDiaANumero(ByVal nombreDia As String) As Integer
    Select Case LCase(Trim(nombreDia))
        Case "lunes", "l", "1": ConvertirDiaANumero = 1
        Case "martes", "m", "2": ConvertirDiaANumero = 2
        Case "miércoles", "miercoles", "x", "3": ConvertirDiaANumero = 3
        Case "jueves", "j", "4": ConvertirDiaANumero = 4
        Case "viernes", "v", "5": ConvertirDiaANumero = 5
        Case "sábado", "sabado", "s", "6": ConvertirDiaANumero = 6
        Case "domingo", "d", "7": ConvertirDiaANumero = 7
        Case Else: ConvertirDiaANumero = 0
    End Select
End Function

Private Function ObtenerLunesDeLaSemana(ByVal fecha As Date) As Date
    ' Esta función obtiene el lunes de la semana para cualquier fecha dada
    Dim diaDeLaSemana As Integer
    
    ' Obtener el día de la semana (1=Lunes, ..., 7=Domingo)
    diaDeLaSemana = Weekday(fecha, vbMonday)
    
    ' Restar los días necesarios para llegar al lunes
    ObtenerLunesDeLaSemana = DateAdd("d", -(diaDeLaSemana - 1), fecha)
End Function

Private Function CalcularFecha(ByVal fechaBase As Date, ByVal diaSemana As String) As Date
    Dim numeroDiaObjetivo As Integer
    Dim lunesDeLaSemana As Date
    
    ' Convertir el nombre del día a número (1=Lunes ... 7=Domingo)
    numeroDiaObjetivo = ConvertirDiaANumero(diaSemana)
    If numeroDiaObjetivo = 0 Then
        CalcularFecha = DateSerial(1900, 1, 1) ' Fecha inválida
        Exit Function
    End If
    
    ' Obtener el lunes de la semana
    lunesDeLaSemana = ObtenerLunesDeLaSemana(fechaBase)
    
    ' Calcular la fecha sumando los días desde el lunes
    CalcularFecha = DateAdd("d", numeroDiaObjetivo - 1, lunesDeLaSemana)
End Function

Private Function ProcesarRegistro(ByRef wsRegistro As Worksheet, _
                                ByRef wsBaseDatos As Worksheet, _
                                ByVal filaRegistro As Long, _
                                ByVal filaBaseDatos As Long, _
                                ByVal fechaBase As Date) As Boolean
                                
    Dim diaSemanaOriginal As String
    Dim fechaCalculada As Date
    
    ' Por defecto, asumimos que el proceso fallará
    ProcesarRegistro = False
    
    ' Obtener el día de la semana original
    diaSemanaOriginal = wsRegistro.Cells(filaRegistro, 2).Value
    
    ' Si está vacío, salimos
    If Trim(diaSemanaOriginal) = "" Then Exit Function
    
    ' Calcular la fecha correspondiente
    fechaCalculada = CalcularFecha(fechaBase, diaSemanaOriginal)
    If fechaCalculada = DateSerial(1900, 1, 1) Then ' Fecha inválida
        MsgBox "El día """ & diaSemanaOriginal & """ en la fila " & filaRegistro & " no es válido." & vbCrLf & _
               "Use nombres de días (ej: Lunes, Martes, etc.)", vbCritical
        Exit Function
    End If
    
    ' Copiar datos
    With wsBaseDatos
        .Cells(filaBaseDatos, 1).Value = fechaCalculada           ' Fecha
        .Cells(filaBaseDatos, 2).Value = wsRegistro.Cells(filaRegistro, 3).Value  ' Número de pedido
        .Cells(filaBaseDatos, 3).Value = diaSemanaOriginal        ' Día de la semana (texto original)
        .Cells(filaBaseDatos, 11).Value = wsRegistro.Cells(filaRegistro, 4).Value ' Metros
        .Cells(filaBaseDatos, 18).Value = wsRegistro.Cells(filaRegistro, 5).Value ' Bandas
        .Cells(filaBaseDatos, 7).Value = wsRegistro.Cells(5, 3).Value            ' Turno
        .Cells(filaBaseDatos, 8).Value = wsRegistro.Cells(6, 3).Value            ' Operario
        
        ' Refilado
        .Cells(filaBaseDatos, 14).Value = IIf(wsRegistro.Cells(filaRegistro, 6).Value = "X", _
                                             "Con Refilado", "Sin Refilado")
        
        ' Tipo
        Select Case wsRegistro.Cells(filaRegistro, 7).Value
            Case "M": .Cells(filaBaseDatos, 15).Value = "Monolámina"
            Case "B": .Cells(filaBaseDatos, 15).Value = "Bicapa"
            Case "T": .Cells(filaBaseDatos, 15).Value = "Tricapa"
            Case Else: .Cells(filaBaseDatos, 15).Value = ""
        End Select
        
        ' Barras
        .Cells(filaBaseDatos, 16).Value = wsRegistro.Cells(filaRegistro, 8).Value
        
        ' Micro
        .Cells(filaBaseDatos, 17).Value = IIf(wsRegistro.Cells(filaRegistro, 9).Value = "X", _
                                             "Con Micro", "Sin Micro")
    End With
    
    ' Indicar que el proceso fue exitoso
    ProcesarRegistro = True
End Function

' Procedimiento para guardar el estado original de una fila para poder deshacerlo después
Private Sub GuardarEstadoFila(ws As Worksheet, fila As Long)
    On Error Resume Next
    
    Dim datos As New Collection
    Dim i As Long
    
    ' Guardar los valores actuales de la fila (todas las columnas relevantes)
    For i = 1 To 20 ' Ajustar según el número real de columnas en uso
        datos.Add ws.Cells(fila, i).Value, CStr(i)
    Next i
    
    ' Añadir a la colección con una clave única
    filasModificadas.Add datos, "Fila_" & fila
    
    On Error GoTo 0
End Sub

' Procedimiento para deshacer los cambios
Private Sub DeshacerCambios(ws As Worksheet)
    On Error Resume Next
    
    Dim filaClave As Variant
    Dim datos As Collection
    Dim fila As Long
    Dim i As Long
    
    ' Recorrer todas las filas guardadas
    For Each filaClave In filasModificadas
        ' Extraer el número de fila de la clave
        fila = CLng(Replace(filaClave, "Fila_", ""))
        ' Obtener los datos originales
        Set datos = filasModificadas(filaClave)
        
        ' Restaurar los valores originales
        For i = 1 To datos.Count
            ws.Cells(fila, i).Value = datos(CStr(i))
        Next i
    Next filaClave
    
    ' Limpiar la colección
    Set filasModificadas = New Collection
    
    On Error GoTo 0
End Sub

Private Sub CommandButton1_Click()
    Dim wsRegistro As Worksheet
    Dim wsBaseDatos As Worksheet
    Dim filaRegistro As Long
    Dim filaBaseDatos As Long
    Dim numPedido As String
    Dim existePedido As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim ultimaFilaRegistro As Long
    Dim fechaBase As Date
    Dim registrosProcesados As Long
    Dim errorOcurrido As Boolean
    
    On Error GoTo ManejadorErrores
    
    ' Definir hojas
    Set wsRegistro = ThisWorkbook.Sheets("Registro")
    Set wsBaseDatos = ThisWorkbook.Sheets("Base de datos")
    
    ' Inicializar colección para deshacer cambios
    Set filasModificadas = New Collection
    
    ' Verificar y obtener la fecha base
    If Not IsDate(wsRegistro.Range("E5").Value) Then
        MsgBox "La fecha en la celda E5 no es válida.", vbCritical
        Exit Sub
    End If
    fechaBase = CDate(wsRegistro.Range("E5").Value)
    
    ' Obtener última fila
    ultimaFilaRegistro = wsRegistro.Cells(Rows.Count, 3).End(xlUp).Row
    registrosProcesados = 0
    errorOcurrido = False
    
    Application.ScreenUpdating = False
    
    ' Procesar registros
    For filaRegistro = 10 To ultimaFilaRegistro
        numPedido = wsRegistro.Cells(filaRegistro, 3).Value
        If numPedido <> "" Then
            ' Verificar si existe el pedido
            existePedido = False
            filaBaseDatos = 0
            
            For filaBaseDatos = 12 To wsBaseDatos.Cells(Rows.Count, 2).End(xlUp).Row
                If wsBaseDatos.Cells(filaBaseDatos, 2).Value = numPedido Then
                    existePedido = True
                    ' Guardar el estado original para poder revertir cambios si es necesario
                    GuardarEstadoFila wsBaseDatos, filaBaseDatos
                    Exit For
                End If
            Next filaBaseDatos
            
            If existePedido Then
                ' Nuevo mensaje con las tres opciones: Editar, Duplicar, Cancelar
                respuesta = MsgBox("El número de pedido " & numPedido & " ya existe. ¿Qué deseas hacer?" & vbCrLf & _
                                 "- Sí: Editar el registro existente" & vbCrLf & _
                                 "- No: Duplicar el registro con nuevos valores" & vbCrLf & _
                                 "- Cancelar: Detener el proceso", _
                                 vbYesNoCancel + vbQuestion, "Pedido Duplicado")
                
                Select Case respuesta
                    Case vbYes
                        ' Editar: Actualizar el registro existente
                        If ProcesarRegistro(wsRegistro, wsBaseDatos, filaRegistro, filaBaseDatos, fechaBase) Then
                            registrosProcesados = registrosProcesados + 1
                        End If
                    
                    Case vbNo
                        ' Duplicar: Crear un nuevo registro
                        filaBaseDatos = wsBaseDatos.Cells(Rows.Count, 2).End(xlUp).Row + 1
                        If ProcesarRegistro(wsRegistro, wsBaseDatos, filaRegistro, filaBaseDatos, fechaBase) Then
                            registrosProcesados = registrosProcesados + 1
                        End If
                    
                    Case vbCancel
                        ' Cancelar: Deshacer los cambios y salir
                        DeshacerCambios wsBaseDatos
                        Application.ScreenUpdating = True
                        MsgBox "Proceso cancelado por el usuario. Los cambios han sido revertidos.", vbInformation
                        Exit Sub
                End Select
            Else
                ' Nuevo registro
                filaBaseDatos = wsBaseDatos.Cells(Rows.Count, 2).End(xlUp).Row + 1
                If ProcesarRegistro(wsRegistro, wsBaseDatos, filaRegistro, filaBaseDatos, fechaBase) Then
                    registrosProcesados = registrosProcesados + 1
                End If
            End If
        End If
    Next filaRegistro
    
    Application.ScreenUpdating = True
    
    ' Mensaje final personalizado
    If registrosProcesados > 0 Then
        MsgBox "Datos guardados correctamente." & vbCrLf & _
               "Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): " & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & _
               "Current User's Login: JhonyAlex" & vbCrLf & _
               "Se procesaron " & registrosProcesados & " registro(s) correctamente.", _
               vbInformation, "Proceso Completado"
    End If
    
    ' Limpiar la colección ya que todo se completó correctamente
    Set filasModificadas = New Collection
    
    Exit Sub

ManejadorErrores:
    errorOcurrido = True
    Application.ScreenUpdating = True
    
    ' Deshacer cambios en caso de error
    DeshacerCambios wsBaseDatos
    
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Los cambios realizados han sido revertidos.", vbCritical, "Error"
    
    Resume Next
End Sub

Private Sub CommandButton2_Click()
    ' Función para limpiar la tabla de registro
    Dim wsRegistro As Worksheet
    Dim respuesta As VbMsgBoxResult
    
    ' Solicitar confirmación al usuario
    respuesta = MsgBox("¿Estás seguro de que deseas limpiar toda la tabla de registro?", _
                     vbYesNo + vbQuestion, "Confirmar Limpieza")
    
    If respuesta = vbNo Then
        ' El usuario canceló la operación
        Exit Sub
    End If
    
    ' Establecer la referencia a la hoja de registro
    Set wsRegistro = ThisWorkbook.Sheets("Registro")
    
    Application.ScreenUpdating = False
    
    ' Mantener la fecha, turno y operario
    Dim fechaActual As Variant
    Dim turnoActual As String
    Dim operarioActual As String
    
    ' Guardar valores actuales
    fechaActual = wsRegistro.Range("E5").Value
    turnoActual = wsRegistro.Range("C5").Value
    operarioActual = wsRegistro.Range("C6").Value
    
    ' Limpiar el rango específico de la tabla: C10:I65
    wsRegistro.Range("C10:I65").ClearContents
    
    ' Restaurar valores de cabecera
    wsRegistro.Range("E5").Value = fechaActual
    wsRegistro.Range("C5").Value = turnoActual
    wsRegistro.Range("C6").Value = operarioActual
    
    Application.ScreenUpdating = True
    
    ' Mostrar mensaje de confirmación con formato personalizado exacto
    MsgBox "Tabla de registro limpiada correctamente." & vbCrLf & _
           "Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): " & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & _
           "Current User's Login: JhonyAlex", _
           vbInformation, "Limpieza Completada"
End Sub
