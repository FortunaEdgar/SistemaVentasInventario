Attribute VB_Name = "modReportes"
'===============================================================
' Modulo: modReportes
' Reportes de ventas e inventario en la hoja "Reportes".
'
'   B3 desde, B4 hasta
'   Fila 7  resumen del periodo (A, B, C, E, F, G)
'   A11     ventas por dia            (A:C)
'   E11     productos mas vendidos    (E:G)
'   I11     productos con stock bajo  (I:K)
'   M11     ultimos movimientos       (M:R)
'===============================================================
Option Explicit

Private Const FILA_DATOS As Long = 11
Private Const ULTIMA_FILA As Long = 1000

'---------------------------------------------------------------
' Boton "Generar reporte"
'---------------------------------------------------------------
Public Sub GenerarReportes()

    Dim hoja As Worksheet
    Dim conexion As ADODB.Connection
    Dim desde As Date
    Dim hasta As Date
    Dim fila As Long

    On Error GoTo ManejarError

    Set hoja = ThisWorkbook.Worksheets(HOJA_REPORTES)

    ' Si no hay fechas se usa el mes actual
    If IsDate(hoja.Range("B3").Value) Then
        desde = CDate(hoja.Range("B3").Value)
    Else
        desde = DateSerial(Year(Date), Month(Date), 1)
        hoja.Range("B3").Value = desde
    End If

    If IsDate(hoja.Range("B4").Value) Then
        hasta = CDate(hoja.Range("B4").Value)
    Else
        hasta = Date
        hoja.Range("B4").Value = hasta
    End If

    If desde > hasta Then
        MsgBox "La fecha 'Desde' no puede ser mayor que la fecha 'Hasta'.", _
               vbExclamation, "Reportes"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    Set conexion = AbrirConexion()

    ' Resumen del periodo
    CargarResumen conexion, hoja, desde, hasta

    ' Ventas por dia
    LimpiarBloque hoja, 1, 3
    CopiarConsulta conexion, hoja.Cells(FILA_DATOS, 1), _
        "SELECT dia, cantidad_ventas, total " & _
        "FROM reporte_ventas_por_dia(CAST(? AS DATE), CAST(? AS DATE));", desde, hasta

    ' Productos mas vendidos
    LimpiarBloque hoja, 5, 7
    CopiarConsulta conexion, hoja.Cells(FILA_DATOS, 5), _
        "SELECT producto, cantidad_vendida, ingresos " & _
        "FROM reporte_productos_mas_vendidos(CAST(? AS DATE), CAST(? AS DATE), 10);", desde, hasta

    ' Stock bajo (no depende de las fechas)
    LimpiarBloque hoja, 9, 11
    CopiarConsulta conexion, hoja.Cells(FILA_DATOS, 9), _
        "SELECT nombre, categoria, stock FROM v_stock_bajo;"

    ' Ultimos movimientos de inventario (kardex)
    LimpiarBloque hoja, 13, 18
    CopiarConsulta conexion, hoja.Cells(FILA_DATOS, 13), _
        "SELECT fecha, producto, tipo, cantidad, stock_resultante, referencia " & _
        "FROM v_kardex ORDER BY id_movimiento DESC LIMIT 30;"

    CerrarConexion conexion

    ' Formatos
    fila = FILA_DATOS + 60
    hoja.Range(hoja.Cells(FILA_DATOS, 1), hoja.Cells(fila, 1)).NumberFormat = "dd/mm/yyyy"
    hoja.Range(hoja.Cells(FILA_DATOS, 3), hoja.Cells(fila, 3)).NumberFormat = "#,##0.00"
    hoja.Range(hoja.Cells(FILA_DATOS, 7), hoja.Cells(fila, 7)).NumberFormat = "#,##0.00"
    hoja.Range(hoja.Cells(FILA_DATOS, 13), hoja.Cells(fila, 13)).NumberFormat = "dd/mm/yyyy hh:mm"

    hoja.Range("A2").Value = "Generado: " & Format$(Now, "dd/mm/yyyy hh:mm")

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Exit Sub

ManejarError:
    MostrarError "generar los reportes", Err.Description
    CerrarConexion conexion

End Sub

' Escribe el resumen del periodo en la fila 7
Private Sub CargarResumen(ByVal conexion As ADODB.Connection, ByVal hoja As Worksheet, _
                          ByVal desde As Date, ByVal hasta As Date)

    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset

    Set comando = NuevoComando(conexion, _
        "SELECT cantidad_ventas, subtotal, itbis, total, ticket_promedio, ventas_anuladas " & _
        "FROM reporte_resumen(CAST(? AS DATE), CAST(? AS DATE));")
    ParamTexto comando, "pDesde", Format$(desde, "yyyy-mm-dd"), 10
    ParamTexto comando, "pHasta", Format$(hasta, "yyyy-mm-dd"), 10
    Set registros = comando.Execute

    hoja.Range("A7").Value = registros.Fields("cantidad_ventas").Value
    hoja.Range("B7").Value = registros.Fields("subtotal").Value
    hoja.Range("C7").Value = registros.Fields("itbis").Value
    hoja.Range("E7").Value = registros.Fields("total").Value
    hoja.Range("F7").Value = registros.Fields("ticket_promedio").Value
    hoja.Range("G7").Value = registros.Fields("ventas_anuladas").Value

    registros.Close
    Set registros = Nothing
    Set comando = Nothing

End Sub

' Borra los datos de un bloque de columnas
Private Sub LimpiarBloque(ByVal hoja As Worksheet, ByVal colInicio As Long, _
                          ByVal colFin As Long)

    hoja.Range(hoja.Cells(FILA_DATOS, colInicio), _
               hoja.Cells(ULTIMA_FILA, colFin)).ClearContents

End Sub

'---------------------------------------------------------------
' Ejecuta una consulta (con fechas opcionales) y pega el
' resultado a partir de la celda indicada.
'---------------------------------------------------------------
Private Sub CopiarConsulta(ByVal conexion As ADODB.Connection, ByVal destino As Range, _
                           ByVal sql As String, _
                           Optional ByVal desde As Variant, Optional ByVal hasta As Variant)

    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset

    Set comando = NuevoComando(conexion, sql)

    If Not IsMissing(desde) Then
        ParamTexto comando, "pDesde", Format$(desde, "yyyy-mm-dd"), 10
        ParamTexto comando, "pHasta", Format$(hasta, "yyyy-mm-dd"), 10
    End If

    Set registros = comando.Execute

    If registros.EOF Then
        destino.Value = "(sin datos)"
    Else
        destino.CopyFromRecordset registros
    End If

    registros.Close
    Set registros = Nothing
    Set comando = Nothing

End Sub
