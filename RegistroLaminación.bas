Option Explicit

Private Sub CommandButton1_Click()
    ' Constantes
    Const HOJA_REGISTRO As String = "Registro"
    Const HOJA_BASE_DATOS As String = "Base de datos"
    Const PRIMERA_FILA_DATOS_REGISTRO As Long = 9
    Const ULTIMA_FILA_DATOS_REGISTRO As Long = 48
    Const PRIMERA_FILA_DATOS_BD As Long = 11
    
    ' Variables
    Dim wsRegistro As Worksheet, wsBaseDatos As Worksheet
    Dim ultimaFilaBD As Long, i As Long, fechaBase As Date, fechaDia As Date
    Dim turno As String, operario As String, dia As String, noPedido As String
    Dim metros As Double, bicapa As String, noTejido As String, tricapa As String, antivaho As String, camisa As String
    Dim totalMetrosDia As Variant
    Dim filaVacia As Boolean, filaExistente As Long
    Dim respuesta As VbMsgBoxResult
    Dim hayPedidosPorDia As Boolean
    Dim registrosProcesados As Long
    
    ' Variables para rastrear cambios
    Dim filasModificadas As New Collection
    
    ' Manejo de Errores
    On Error GoTo ManejadorErrores
    
    ' Referencias a las Hojas
    Set wsRegistro = ThisWorkbook.Worksheets(HOJA_REGISTRO)
    Set wsBaseDatos = ThisWorkbook.Worksheets(HOJA_BASE_DATOS)
    
    ' Inicializar contador
    registrosProcesados = 0
    
    ' Obtener Turno y Operario
    turno = wsRegistro.Range("C4").Value
    operario = wsRegistro.Range("C5").Value
    
    ' Validación
    If turno = "" Or operario = "" Then
        MsgBox "Por favor, ingresa el Turno y el Operario", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If
    
    ' Validar fecha
    If IsEmpty(wsRegistro.Range("H4").Value) Or Not IsDate(wsRegistro.Range("H4").Value) Then
        MsgBox "Por favor, ingresa una fecha válida en la celda H4", vbExclamation, "Fecha Inválida"
        Exit Sub
    End If
    
    ' Obtener la Última Fila en "Base de datos"
    ultimaFilaBD = wsBaseDatos.Cells(wsBaseDatos.Rows.count, "A").End(xlUp).Row
    If ultimaFilaBD < PRIMERA_FILA_DATOS_BD - 1 Then
        ultimaFilaBD = PRIMERA_FILA_DATOS_BD - 1
    End If
    
    ' Calcular fecha base (lunes de la semana)
    fechaBase = wsRegistro.Range("H4").Value - Weekday(wsRegistro.Range("H4").Value, vbMonday) + 1
    
    ' Verificar si hay pedidos
    hayPedidosPorDia = False
    For i = PRIMERA_FILA_DATOS_REGISTRO To ULTIMA_FILA_DATOS_REGISTRO
        If Not IsEmpty(wsRegistro.Cells(i, "C").Value) Then
            hayPedidosPorDia = True
            Exit For
        End If
    Next i
    
    ' Desactivar actualización para mejor rendimiento
    Application.ScreenUpdating = False
    
    ' Ciclo Principal
    For i = PRIMERA_FILA_DATOS_REGISTRO To ULTIMA_FILA_DATOS_REGISTRO
        On Error GoTo 0
        
        ' Verificar si hay NoPedido
        If Not IsEmpty(wsRegistro.Cells(i, "C").Value) Then
            dia = wsRegistro.Cells(i, "B").Value
            noPedido = wsRegistro.Cells(i, "C").Value
            metros = wsRegistro.Cells(i, "D").Value
            camisa = wsRegistro.Cells(i, "E").Value ' Obtener el valor de Camisa
            
            ' Obtener valores de productos
            bicapa = IIf(wsRegistro.Cells(i, "F").Value = "X", "Bicapa", "N/A")
            noTejido = IIf(wsRegistro.Cells(i, "G").Value = "X", "No tejido", "N/A")
            tricapa = IIf(wsRegistro.Cells(i, "H").Value = "X", "Tricapa", "N/A")
            antivaho = IIf(wsRegistro.Cells(i, "I").Value = "X", "Antivaho", "N/A")
            
            ' Verificar si existe pedido
            filaExistente = ExisteNoPedido(wsBaseDatos, noPedido)
            
                       If filaExistente > 0 Then
                ' Mostrar opciones para pedido duplicado
                respuesta = MsgBox("El número de pedido " & noPedido & " ya existe. ¿Qué deseas hacer?" & vbCrLf & _
                                 "- Sí: Editar el registro existente" & vbCrLf & _
                                 "- No: Duplicar el registro con nuevos valores" & vbCrLf & _
                                 "- Cancelar: Detener el proceso", _
                                 vbYesNoCancel + vbQuestion, "Pedido Duplicado")
                
                Select Case respuesta
                    Case vbYes
                        ' Editar
                        ultimaFilaBD = filaExistente
                        ' Guardar estado para deshacer
                        GuardarEstadoFila wsBaseDatos, ultimaFilaBD, filasModificadas
                        
                    Case vbNo
                        ' Duplicar
                        ultimaFilaBD = ultimaFilaBD + 1
                        
                    Case vbCancel
                        ' Cancelar y mantener los datos en el formulario
                        DeshacerCambios wsBaseDatos, filasModificadas
                        Application.ScreenUpdating = True
                        MsgBox "Proceso cancelado por el usuario. Los datos se mantienen en el formulario.", vbInformation
                        Exit Sub
                End Select
            End If
            
            ' Calcular Fecha según el día
            Select Case dia
                Case "Lunes"
                    fechaDia = fechaBase
                    totalMetrosDia = IIf(hayPedidosPorDia, wsRegistro.Range("J9").Value, "")
                Case "Martes"
                    fechaDia = fechaBase + 1
                    totalMetrosDia = IIf(hayPedidosPorDia, wsRegistro.Range("J17").Value, "")
                Case "Miércoles"
                    fechaDia = fechaBase + 2
                    totalMetrosDia = IIf(hayPedidosPorDia, wsRegistro.Range("J25").Value, "")
                Case "Jueves"
                    fechaDia = fechaBase + 3
                    totalMetrosDia = IIf(hayPedidosPorDia, wsRegistro.Range("J33").Value, "")
                Case "Viernes"
                    fechaDia = fechaBase + 4
                    totalMetrosDia = IIf(hayPedidosPorDia, wsRegistro.Range("J41").Value, "")
                Case "Sábado"
                    fechaDia = fechaBase + 5
                    totalMetrosDia = "" ' No hay valor específico para Sábado
                Case "Domingo"
                    fechaDia = fechaBase + 6
                    totalMetrosDia = "" ' No hay valor específico para Domingo
                Case Else
                    MsgBox "Día no reconocido en la fila " & i & ": " & dia, vbExclamation
                    GoTo SiguienteRegistro
            End Select
            
            ' Escribir datos
            With wsBaseDatos
                .Cells(ultimaFilaBD, "A").Value = fechaDia
                .Cells(ultimaFilaBD, "B").Value = WorksheetFunction.IsoWeekNum(fechaDia)
                .Cells(ultimaFilaBD, "C").Value = dia
                .Cells(ultimaFilaBD, "D").Value = turno
                .Cells(ultimaFilaBD, "E").Value = operario
                ' Columna F (Máquina) se deja en blanco
                .Cells(ultimaFilaBD, "G").Value = noPedido
                .Cells(ultimaFilaBD, "H").Value = ContarPedidosPorDia(wsRegistro, dia)
                .Cells(ultimaFilaBD, "I").Value = metros
                ' Columna J (Metros por día) se deja en blanco
                .Cells(ultimaFilaBD, "K").Value = bicapa
                .Cells(ultimaFilaBD, "L").Value = noTejido
                .Cells(ultimaFilaBD, "M").Value = tricapa
                .Cells(ultimaFilaBD, "N").Value = antivaho
                .Cells(ultimaFilaBD, "O").Value = camisa ' Guardar el valor de Camisa en la columna O
            End With
            
            registrosProcesados = registrosProcesados + 1
        End If
        
SiguienteRegistro:
    Next i
    
    Application.ScreenUpdating = True    
    If registrosProcesados > 0 Then
        MsgBox "Datos guardados correctamente." & vbCrLf & _
               "Fecha actual: " & Format(Now, "yyyy-MM-dd HH:mm:ss") & vbCrLf & _
               "Current User's Login: JhonyAlvarez", vbInformation, "Éxito"
        
        ' Llamar a la función de limpieza
        LimpiarDespuesDeEnvio
    End If
    Exit Sub
    
ManejadorErrores:
    Application.ScreenUpdating = True
    DeshacerCambios wsBaseDatos, filasModificadas
    MsgBox "Ocurrió un error: " & Err.Description & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "En la fila (i): " & i & vbCrLf & _
           "Los cambios realizados han sido revertidos.", vbCritical, "Error"
End Sub

' Función para verificar si un número de pedido ya existe
Private Function ExisteNoPedido(ws As Worksheet, noPedido As String) As Long
    Dim ultimaFila As Long
    Dim i As Long
    
    ultimaFila = ws.Cells(ws.Rows.count, "G").End(xlUp).Row
    
    For i = 11 To ultimaFila
        If ws.Cells(i, "G").Value = noPedido Then
            ExisteNoPedido = i
            Exit Function
        End If
    Next i
    
    ExisteNoPedido = 0
End Function

' Función para contar pedidos por día en la hoja de registro según rangos específicos
Private Function ContarPedidosPorDia(ws As Worksheet, dia As String) As Long
    Dim count As Long
    Dim filaInicio As Long, filaFin As Long
    Dim i As Long
    
    ' Determinar el rango correcto según el día
    Select Case dia
        Case "Lunes"
            filaInicio = 9
            filaFin = 16
        Case "Martes"
            filaInicio = 17
            filaFin = 24
        Case "Miércoles"
            filaInicio = 25
            filaFin = 32
        Case "Jueves"
            filaInicio = 33
            filaFin = 40
        Case "Viernes"
            filaInicio = 41
            filaFin = 48
        Case Else
            ' Para otros días que no tienen un rango específico
            count = 0
            Exit Function
    End Select
    
    count = 0
    
    ' Contar solo dentro del rango correspondiente al día
    For i = filaInicio To filaFin
        If Not IsEmpty(ws.Cells(i, "C").Value) Then
            count = count + 1
        End If
    Next i
    
    ContarPedidosPorDia = count
End Function

' Guardar estado de fila para deshacer
Private Sub GuardarEstadoFila(ws As Worksheet, fila As Long, ByRef coleccion As Collection)
    Dim datos(1 To 15) As Variant
    Dim i As Long
    
    ' Guardar valores actuales
    For i = 1 To 15
        datos(i) = ws.Cells(fila, i).Value
    Next i
    
    ' Añadir a la colección
    On Error Resume Next
    coleccion.Add datos, "Fila_" & fila
    On Error GoTo 0
End Sub

' Deshacer cambios
Private Sub DeshacerCambios(ws As Worksheet, ByRef coleccion As Collection)
    Dim fila As Long
    Dim datos As Variant
    Dim i As Integer, j As Integer
    
    ' Recorrer filas guardadas
    For i = 1 To coleccion.count
        On Error Resume Next
        datos = coleccion.Item(i)
        fila = CInt(Replace(coleccion(i).key, "Fila_", ""))
        
        ' Restaurar valores
        For j = 1 To 15
            ws.Cells(fila, j).Value = datos(j)
        Next j
        On Error GoTo 0
    Next i
End Sub

' Agregar botón para la fecha actual
Private Sub CommandButton2_Click()
    ' Obtener la fecha actual
    Dim fechaActual As Date
    fechaActual = Date
    
    ' Escribir la fecha formateada en H4
    With ThisWorkbook.Sheets("Registro")
        .Range("H4").Value = Format(fechaActual, "dd/mm/yyyy")
    End With
End Sub

' Agregar botón para limpiar el formulario
Private Sub CommandButton3_Click()
    ' Especifica la hoja de trabajo.
    With ThisWorkbook.Sheets("Registro")
        ' Borra el contenido de los rangos especificados.
        .Range("C4:C5").ClearContents ' Turno y Operario
        .Range("C9:I48").ClearContents ' Datos de la tabla
        .Range("H4").ClearContents ' Fecha
    End With
    
    ' Mensaje opcional
    ' MsgBox "El contenido de las celdas ha sido borrado.", vbInformation, "Borrado Exitoso"
End Sub

' Nueva función para limpiar después del envío
Private Sub LimpiarDespuesDeEnvio()
    With ThisWorkbook.Sheets("Registro")
        ' Limpiar todo: datos de la tabla, turno, operario y fecha
        .Range("C9:I48").ClearContents  ' Datos de la tabla
        .Range("C4:C5").ClearContents   ' Turno y Operario
        .Range("H4").ClearContents      ' Fecha
    End With
End Sub
