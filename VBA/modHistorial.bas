Attribute VB_Name = "modHistorial"
'===============================================================
' Modulo: modHistorial
' Historial de ventas y anulacion de facturas.
'
' Distribucion de la hoja "Historial":
'   B3 desde, D3 hasta, F3 cliente (filtros opcionales)
'   Fila 5 encabezados, datos desde la fila 6 (columnas A:I)
'===============================================================
Option Explicit

Private Const FILA_DATOS As Long = 6
Private Const ULTIMA_FILA As Long = 5000

'---------------------------------------------------------------
' Boton "Actualizar": carga las ventas segun los filtros
'---------------------------------------------------------------
Public Sub CargarHistorial()

    Dim hoja As Worksheet
    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset
    Dim desde As Date
    Dim hasta As Date
    Dim cliente As String
    Dim filas As Long

    On Error GoTo ManejarError

    Set hoja = ThisWorkbook.Worksheets(HOJA_HISTORIAL)

    ' Filtros (si estan vacios se muestran todas las ventas)
    If IsDate(hoja.Range("B3").Value) Then
        desde = CDate(hoja.Range("B3").Value)
    Else
        desde = DateSerial(2000, 1, 1)
    End If

    If IsDate(hoja.Range("D3").Value) Then
        hasta = CDate(hoja.Range("D3").Value)
    Else
        hasta = DateSerial(2100, 12, 31)
    End If

    If desde > hasta Then
        MsgBox "La fecha 'Desde' no puede ser mayor que la fecha 'Hasta'.", _
               vbExclamation, "Historial"
        Exit Sub
    End If

    cliente = TextoCelda(hoja.Range("F3"))

    Application.ScreenUpdating = False

    Set conexion = AbrirConexion()
    Set comando = NuevoComando(conexion, _
        "SELECT id_venta, numero, fecha, cliente, articulos, subtotal, itbis, total, estado " & _
        "FROM v_ventas_resumen " & _
        "WHERE fecha >= CAST(? AS DATE) " & _
        "  AND fecha < CAST(? AS DATE) + 1 " & _
        "  AND cliente ILIKE CAST(? AS VARCHAR) " & _
        "ORDER BY id_venta DESC;")

    ParamTexto comando, "pDesde", Format$(desde, "yyyy-mm-dd"), 10
    ParamTexto comando, "pHasta", Format$(hasta, "yyyy-mm-dd"), 10
    ParamTexto comando, "pCliente", "%" & cliente & "%", 110

    Set registros = comando.Execute

    hoja.Range(hoja.Cells(FILA_DATOS, 1), hoja.Cells(ULTIMA_FILA, 9)).ClearContents

    If Not registros.EOF Then
        hoja.Cells(FILA_DATOS, 1).CopyFromRecordset registros
    End If

    registros.Close
    Set registros = Nothing
    CerrarConexion conexion

    ' Formatos de las columnas
    filas = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row

    If filas >= FILA_DATOS Then
        hoja.Range(hoja.Cells(FILA_DATOS, 3), hoja.Cells(filas, 3)).NumberFormat = "dd/mm/yyyy hh:mm"
        hoja.Range(hoja.Cells(FILA_DATOS, 6), hoja.Cells(filas, 8)).NumberFormat = "#,##0.00"
    End If

    Application.ScreenUpdating = True
    Exit Sub

ManejarError:
    MostrarError "cargar el historial de ventas", Err.Description
    CerrarConexion conexion

End Sub

' Id de la venta seleccionada en el historial (0 si no hay)
Private Function VentaSeleccionada() As Long

    Dim valor As Variant

    VentaSeleccionada = 0

    If ActiveSheet.Name <> HOJA_HISTORIAL Then Exit Function
    If ActiveCell.Row < FILA_DATOS Then Exit Function

    valor = ActiveSheet.Cells(ActiveCell.Row, 1).Value
    If EsEntero(valor) Then VentaSeleccionada = CLng(valor)

End Function

'---------------------------------------------------------------
' Boton "Ver factura": abre la factura de la fila seleccionada
'---------------------------------------------------------------
Public Sub VerFacturaSeleccionada()

    Dim idVenta As Long

    idVenta = VentaSeleccionada()

    If idVenta = 0 Then
        BuscarFactura
    Else
        MostrarFactura idVenta
    End If

End Sub

'---------------------------------------------------------------
' Boton "Anular venta": anula la venta de la fila seleccionada
'---------------------------------------------------------------
Public Sub AnularVentaSeleccionada()

    Dim idVenta As Long
    Dim texto As String
    Dim cancelado As Boolean

    idVenta = VentaSeleccionada()

    If idVenta = 0 Then
        texto = PedirTexto("Seleccione una venta en la tabla o escriba el numero" & vbCrLf & _
                           "de la factura que desea anular:", "Anular venta", cancelado)
        If cancelado Or texto = "" Then Exit Sub
        idVenta = IdDesdeNumeroFactura(texto)
        If idVenta <= 0 Then
            MsgBox "Numero de factura no valido.", vbExclamation, "Anular venta"
            Exit Sub
        End If
    End If

    If AnularVentaPorId(idVenta) Then CargarHistorial

End Sub

'---------------------------------------------------------------
' Anula una venta en PostgreSQL. Devuelve True si se anulo.
' La funcion anular_venta() devuelve los productos al inventario.
'---------------------------------------------------------------
Public Function AnularVentaPorId(ByVal idVenta As Long) As Boolean

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command

    AnularVentaPorId = False

    If MsgBox("¿Desea ANULAR la factura " & NumeroFactura(idVenta) & "?" & vbCrLf & vbCrLf & _
              "Los productos se devolveran al inventario y la venta" & vbCrLf & _
              "dejara de contar en los reportes. Esta accion no se puede deshacer.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Anular venta") <> vbYes Then
        Exit Function
    End If

    On Error GoTo ManejarError

    Set conexion = AbrirConexion()
    Set comando = NuevoComando(conexion, "SELECT anular_venta(?);")
    ParamEntero comando, "pId", idVenta
    comando.Execute
    CerrarConexion conexion

    Application.ScreenUpdating = False
    RefrescarTabla HOJA_PRODUCTOS, TABLA_PRODUCTOS, "productos"
    RefrescarTabla HOJA_VENTAS, TABLA_VENTAS, "ventas"
    ActualizarIndicadores
    Application.ScreenUpdating = True

    MsgBox "La factura " & NumeroFactura(idVenta) & " fue anulada.", vbInformation, "Anular venta"
    AnularVentaPorId = True
    Exit Function

ManejarError:
    MostrarError "anular la venta", Err.Description
    CerrarConexion conexion

End Function
