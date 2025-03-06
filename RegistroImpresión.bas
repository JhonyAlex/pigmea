Option Explicit

Private Sub CommandButton1_Click()

  ' Especifica la hoja de trabajo.
  With ThisWorkbook.Sheets("Registro")

    ' Borra el contenido de los rangos especificados.
    .Range("C7:G13").ClearContents
    .Range("C3").ClearContents
    .Range("C4").ClearContents
    .Range("F3").ClearContents

    'Opcional: Un mensaje para indicar que se completó la acción.  Es buena práctica dar feedback al usuario.
    'MsgBox "El contenido de las celdas ha sido borrado.", vbInformation, "Borrado Exitoso"

  End With

End Sub

Private Sub CommandButton2_Click()

    ' --- Constantes ---
    Const HOJA_REGISTRO As String = "Registro"
    Const HOJA_BASE_DATOS As String = "Base de datos"
    Const HOJA_DATOS As String = "Datos"
    Const PRIMERA_FILA_DATOS_REGISTRO As Long = 7
    Const ULTIMA_FILA_DATOS_REGISTRO As Long = 13
    Const PRIMERA_FILA_DATOS_BD As Long = 23

    ' --- Variables ---
    Dim wsRegistro As Worksheet
    Dim wsBaseDatos As Worksheet
    Dim wsDatos As Worksheet
    Dim ultimaFilaBD As Long
    Dim i As Long
    Dim fechaBase As Date  ' <-- Fecha base para todos los cálculos
    Dim fechaDia As Date
    Dim semanaActual As Integer
    Dim turno As String
    Dim operario As String
    Dim dia As String
    Dim totalPedidos As Long
    Dim cambiosTinta As Long
    Dim camisas As Long
    Dim totalMetros As Double
    Dim cambiosTransparencia As String
    Dim filaVacia As Boolean
    Dim acumuladoSemanal As Double
    Dim fechaInicialRegistro As Date  ' <-- Fecha del primer registro no vacío

    ' --- Manejo de Errores ---
    On Error GoTo ManejadorErrores

    ' --- Referencias a las Hojas ---
    Set wsRegistro = ThisWorkbook.Worksheets(HOJA_REGISTRO)
    Set wsBaseDatos = ThisWorkbook.Worksheets(HOJA_BASE_DATOS)
    Set wsDatos = ThisWorkbook.Worksheets(HOJA_DATOS)

    ' --- Obtener Turno y Operario ---
    turno = wsRegistro.Range("C3").Value
    operario = wsRegistro.Range("C4").Value

    ' --- Validación (solo Turno y Operario) ---
    If turno = "" Or operario = "" Then
        MsgBox "Por favor, ingresa el Turno y el Operario", vbExclamation, "Datos Incompletos"
        Exit Sub
    End If

    ' --- Obtener la Última Fila en "Base de Datos" ---
    ultimaFilaBD = wsBaseDatos.Cells(Rows.Count, "A").End(xlUp).Row
    If ultimaFilaBD < PRIMERA_FILA_DATOS_BD - 1 Then
        ultimaFilaBD = PRIMERA_FILA_DATOS_BD - 1
    End If
    ultimaFilaBD = ultimaFilaBD + 1

    ' --- Encontrar la Fecha Inicial del Registro (Primer registro con datos) ---
    fechaInicialRegistro = 0  ' Inicializar en un valor que indique que no se ha encontrado
    For i = PRIMERA_FILA_DATOS_REGISTRO To ULTIMA_FILA_DATOS_REGISTRO
        filaVacia = (wsRegistro.Cells(i, "C").Value = "" And _
                     wsRegistro.Cells(i, "D").Value = "" And _
                     wsRegistro.Cells(i, "E").Value = "" And _
                     wsRegistro.Cells(i, "F").Value = "" And _
                     wsRegistro.Cells(i, "G").Value = "")
        If Not filaVacia Then
           
           If wsRegistro.Cells(i, "B").Value = "Lunes 1" Then
               fechaInicialRegistro = wsRegistro.Range("F3").Value + 7
            ElseIf wsRegistro.Cells(i, "B").Value = "Martes 2" Then
                fechaInicialRegistro = wsRegistro.Range("F3").Value + 8
           Else
            fechaInicialRegistro = wsRegistro.Range("F3").Value
            End If

            Exit For  ' Salir del bucle una vez encontrada la primera fecha
        End If
    Next i
     ' --- Si NO se encontró una fecha en el rango, usar F3 ---
    If fechaInicialRegistro = 0 Then
        fechaInicialRegistro = wsRegistro.Range("F3").Value
      If IsEmpty(fechaInicialRegistro) Or Not IsDate(fechaInicialRegistro) Then
        MsgBox "No se encontró una fecha válida en el rango de C7:C13 ni en F3. Se usará la fecha actual.", vbExclamation
        fechaInicialRegistro = Date
      End If
    End If


    ' --- Ajustar fechaInicialRegistro al Lunes de esa semana ---
    fechaBase = fechaInicialRegistro - Weekday(fechaInicialRegistro, vbMonday) + 1

    ' --- Ciclo Principal ---
    acumuladoSemanal = 0  ' Inicializar el acumulado

    For i = PRIMERA_FILA_DATOS_REGISTRO To ULTIMA_FILA_DATOS_REGISTRO

        filaVacia = (wsRegistro.Cells(i, "C").Value = "" And _
                    wsRegistro.Cells(i, "D").Value = "" And _
                    wsRegistro.Cells(i, "E").Value = "" And _
                    wsRegistro.Cells(i, "F").Value = "" And _
                    wsRegistro.Cells(i, "G").Value = "")

        If Not filaVacia Then

            dia = wsRegistro.Cells(i, "B").Value
            totalPedidos = wsRegistro.Cells(i, "C").Value
            camisas = wsRegistro.Cells(i, "D").Value
            cambiosTinta = wsRegistro.Cells(i, "E").Value
            totalMetros = wsRegistro.Cells(i, "F").Value
            cambiosTransparencia = wsRegistro.Cells(i, "G").Value

            ' --- Calcular Fecha ---
            Select Case dia
                Case "Lunes", "Lunes 1"
                    fechaDia = fechaBase
                     If dia = "Lunes 1" Then fechaDia = fechaDia + 7
                Case "Martes", "Martes 2"
                    fechaDia = fechaBase + 1
                    If dia = "Martes 2" Then fechaDia = fechaDia + 7
                Case "Miércoles"
                    fechaDia = fechaBase + 2
                Case "Jueves"
                    fechaDia = fechaBase + 3
                Case "Viernes"
                    fechaDia = fechaBase + 4
                Case "Sábado"
                    fechaDia = fechaBase + 5
                Case "Domingo"
                    fechaDia = fechaBase + 6
                 Case Else
                    Err.Raise vbObjectError + 513, "CommandButton2_Click", "Día de la semana inválido: " & dia
            End Select

            ' --- Escribir Datos Básicos ---
            wsBaseDatos.Cells(ultimaFilaBD, "A").Value = fechaDia
            wsBaseDatos.Cells(ultimaFilaBD, "B").Value = WorksheetFunction.IsoWeekNum(fechaDia)
            wsBaseDatos.Cells(ultimaFilaBD, "C").Value = dia
            wsBaseDatos.Cells(ultimaFilaBD, "D").Value = turno
            wsBaseDatos.Cells(ultimaFilaBD, "E").Value = operario

             ' --- Verificar si la hoja "Datos" tiene datos ---
            If WorksheetFunction.CountA(wsDatos.Range("C:D")) = 0 Then
                MsgBox "La hoja 'Datos' está vacía o no tiene datos en las columnas C y D.", vbCritical, "Error en Datos"
                GoTo ManejadorErrores
            End If

            ' --- Construir la fórmula INDICE/COINCIDIR ---
            Dim formulaAlternativa As String
            formulaAlternativa = "=IFERROR(INDEX(Datos!$D:$D,MATCH(TEXT(" & wsBaseDatos.Cells(ultimaFilaBD, "E").Address & ",""@""" & "),Datos!$C:$C,0)),""No está"")"
            wsBaseDatos.Cells(ultimaFilaBD, "F").Formula = formulaAlternativa

            wsBaseDatos.Cells(ultimaFilaBD, "G").Value = camisas
            wsBaseDatos.Cells(ultimaFilaBD, "H").Value = cambiosTinta
            wsBaseDatos.Cells(ultimaFilaBD, "I").Value = totalMetros
            wsBaseDatos.Cells(ultimaFilaBD, "J").Value = totalPedidos

            ' --- Cálculo del Acumulado (Columna K) ---
            acumuladoSemanal = acumuladoSemanal + totalMetros
            wsBaseDatos.Cells(ultimaFilaBD, "K").Value = acumuladoSemanal

            wsBaseDatos.Cells(ultimaFilaBD, "L").Value = cambiosTransparencia
            ultimaFilaBD = ultimaFilaBD + 1
        End If
    Next i

    MsgBox "Datos guardados correctamente.", vbInformation, "Éxito"
    Exit Sub

ManejadorErrores:
    MsgBox "Ocurrió un error: " & Err.Description & vbCrLf & _
           "Número de error: " & Err.Number & vbCrLf & _
           "En la fila (i): " & i, vbCritical, "Error"

End Sub




Private Sub CommandButton3_Click()

  ' Obtener la fecha actual
  Dim fechaActual As Date
  fechaActual = Date

  ' Formatear la fecha como "dd/mm/yyyy"
  ' Usar la función Format para controlar la presentación de la fecha.

    With ThisWorkbook.Sheets("Registro")
        .Range("F3").Value = Format(fechaActual, "dd/mm/yyyy")
    End With
  
  'Otras opciones. Comentadas. Descomentar (quitar el ') la que se necesite.
  
  ' Mostrar la fecha en un cuadro de mensaje (opcional):
  ' MsgBox "La fecha actual es: " & Format(fechaActual, "dd/mm/yyyy"), vbInformation, "Fecha Actual"

  ' Escribir la fecha en una celda específica (ejemplo: A1 de la hoja activa):
  ' ActiveSheet.Range("A1").Value = Format(fechaActual, "dd/mm/yyyy")
    
    ' Escribir la fecha en una celda y hoja específica
  'ThisWorkbook.Sheets("NombreDeLaHoja").Range("A1").Value = Format(fechaActual, "dd/mm/yyyy")

    'Si el resultado se desea ver en el "Immediate Window" (Ctrl + G)
    'Debug.Print Format(fechaActual, "dd/mm/yyyy")

End Sub
