Attribute VB_Name = "modVentas"
'===============================================================
' Modulo: modVentas
' Registro de ventas desde la hoja "Nueva Venta".
'
' Distribucion de la hoja:
'   B3        -> nombre del cliente (lista desplegable)
'   A6:A26    -> productos (lista desplegable)
'   B6:B26    -> cantidades
'   C, D, E   -> precio, subtotal y stock (formulas)
'   F6:F26    -> id del producto (formula)
'   I3:I5     -> subtotal sin ITBIS, ITBIS incluido y total
'               (los precios ya incluyen el ITBIS)
'   I6        -> id del cliente (formula)
'===============================================================
Option Explicit

Private Const FILA_INICIO As Long = 6
Private Const FILA_FIN As Long = 26

Private Const COL_PRODUCTO As Long = 1
Private Const COL_CANTIDAD As Long = 2
Private Const COL_STOCK As Long = 5
Private Const COL_ID As Long = 6

'---------------------------------------------------------------
' Registra la venta que esta en la hoja "Nueva Venta"
'---------------------------------------------------------------
Public Sub RegistrarVenta()

    Dim hoja As Worksheet
    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim resultado As ADODB.Recordset

    Dim cliente As String
    Dim idCliente As Long
    Dim idVenta As Long
    Dim fila As Long
    Dim i As Long
    Dim lineas As Long
    Dim producto As String
    Dim cantidad As Variant
    Dim stock As Variant
    Dim idProducto As Variant
    Dim ids() As String
    Dim cantidades() As String
    Dim resumen As String

    Const TITULO As String = "Registrar venta"

    Set hoja = ThisWorkbook.Worksheets(HOJA_VENTA)
    hoja.Activate

    '--- 1. Validar el cliente -----------------------------------
    cliente = TextoCelda(hoja.Range("B3"))

    If cliente = "" Then
        MsgBox "Seleccione el cliente de la venta.", vbExclamation, TITULO
        hoja.Activate
        hoja.Range("B3").Select
        Exit Sub
    End If

    If Not EsEntero(hoja.Range("I6").Value) Then
        MsgBox "El cliente """ & cliente & """ no existe en el sistema." & vbCrLf & _
               "Seleccionelo de la lista o registrelo primero.", vbExclamation, TITULO
        Exit Sub
    End If

    idCliente = CLng(hoja.Range("I6").Value)

    '--- 2. Leer y validar las lineas ----------------------------
    ReDim ids(1 To FILA_FIN - FILA_INICIO + 1)
    ReDim cantidades(1 To FILA_FIN - FILA_INICIO + 1)

    For fila = FILA_INICIO To FILA_FIN

        producto = TextoCelda(hoja.Cells(fila, COL_PRODUCTO))

        If producto <> "" Then

            idProducto = hoja.Cells(fila, COL_ID).Value
            cantidad = hoja.Cells(fila, COL_CANTIDAD).Value
            stock = hoja.Cells(fila, COL_STOCK).Value

            If Not EsEntero(idProducto) Then
                MsgBox "El producto """ & producto & """ (fila " & fila & _
                       ") no existe en el catalogo.", vbExclamation, TITULO
                hoja.Cells(fila, COL_PRODUCTO).Select
                Exit Sub
            End If

            If Not EsEntero(cantidad) Then
                MsgBox "Escriba una cantidad valida para """ & producto & _
                       """ (fila " & fila & ").", vbExclamation, TITULO
                hoja.Cells(fila, COL_CANTIDAD).Select
                Exit Sub
            End If

            If CLng(cantidad) < 1 Then
                MsgBox "La cantidad de """ & producto & """ debe ser mayor que cero.", _
                       vbExclamation, TITULO
                hoja.Cells(fila, COL_CANTIDAD).Select
                Exit Sub
            End If

            If EsEntero(stock) Then
                If CLng(cantidad) > CLng(stock) Then
                    MsgBox "No hay suficiente stock de """ & producto & """." & vbCrLf & _
                           "Disponible: " & stock & "   Solicitado: " & cantidad, _
                           vbExclamation, TITULO
                    hoja.Cells(fila, COL_CANTIDAD).Select
                    Exit Sub
                End If
            End If

            ' Evitar el mismo producto en dos lineas
            For i = 1 To lineas
                If ids(i) = CStr(CLng(idProducto)) Then
                    MsgBox "El producto """ & producto & """ esta repetido." & vbCrLf & _
                           "Use una sola linea y sume las cantidades.", vbExclamation, TITULO
                    hoja.Cells(fila, COL_PRODUCTO).Select
                    Exit Sub
                End If
            Next i

            lineas = lineas + 1
            ids(lineas) = CStr(CLng(idProducto))
            cantidades(lineas) = CStr(CLng(cantidad))

        ElseIf TextoCelda(hoja.Cells(fila, COL_CANTIDAD)) <> "" Then

            MsgBox "La fila " & fila & " tiene cantidad pero no tiene producto.", _
                   vbExclamation, TITULO
            hoja.Cells(fila, COL_PRODUCTO).Select
            Exit Sub

        End If

    Next fila

    If lineas = 0 Then
        MsgBox "Agregue al menos un producto a la venta.", vbExclamation, TITULO
        hoja.Cells(FILA_INICIO, COL_PRODUCTO).Select
        Exit Sub
    End If

    ReDim Preserve ids(1 To lineas)
    ReDim Preserve cantidades(1 To lineas)

    '--- 3. Confirmar --------------------------------------------
    resumen = "Cliente: " & cliente & vbCrLf & _
              "Productos: " & lineas & vbCrLf & vbCrLf & _
              "Subtotal: " & Format$(hoja.Range("I3").Value, FORMATO_MONEDA) & vbCrLf & _
              "ITBIS incluido: " & Format$(hoja.Range("I4").Value, FORMATO_MONEDA) & vbCrLf & _
              "TOTAL: " & Format$(hoja.Range("I5").Value, FORMATO_MONEDA) & vbCrLf & vbCrLf & _
              "¿Desea registrar la venta?"

    If MsgBox(resumen, vbQuestion + vbYesNo, TITULO) <> vbYes Then Exit Sub

    '--- 4. Guardar en PostgreSQL --------------------------------
    ' La funcion registrar_venta() guarda el encabezado, el detalle,
    ' descuenta el stock y calcula el ITBIS en UNA sola transaccion.
    On Error GoTo ManejarError

    Application.Cursor = xlWait

    Set conexion = AbrirConexion()

    Set comando = NuevoComando(conexion, _
        "SELECT registrar_venta(?, CAST(? AS INTEGER[]), CAST(? AS INTEGER[])) AS id_venta;")

    ParamEntero comando, "pIdCliente", idCliente
    ParamTexto comando, "pProductos", "{" & Join(ids, ",") & "}", 4000
    ParamTexto comando, "pCantidades", "{" & Join(cantidades, ",") & "}", 4000

    Set resultado = comando.Execute
    idVenta = CLng(resultado.Fields("id_venta").Value)

    resultado.Close
    Set resultado = Nothing
    Set comando = Nothing
    CerrarConexion conexion

    '--- 5. Actualizar Excel y mostrar la factura ----------------
    Application.ScreenUpdating = False

    LimpiarFormularioVenta
    RefrescarTabla HOJA_PRODUCTOS, TABLA_PRODUCTOS, "productos"
    RefrescarTabla HOJA_VENTAS, TABLA_VENTAS, "ventas"
    RefrescarTabla HOJA_DETALLE, TABLA_DETALLE, "detalle_venta"
    ActualizarIndicadores

    MostrarFactura idVenta

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

    If MsgBox("Venta registrada correctamente." & vbCrLf & vbCrLf & _
              "Factura: " & NumeroFactura(idVenta) & vbCrLf & vbCrLf & _
              "¿Desea guardar la factura en PDF?", _
              vbInformation + vbYesNo, TITULO) = vbYes Then
        ExportarFacturaPDF
    End If

    Exit Sub

ManejarError:
    MostrarError "registrar la venta", Err.Description
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Boton "Limpiar": borra la venta en pantalla (pide confirmacion)
'---------------------------------------------------------------
Public Sub LimpiarVenta()

    If MsgBox("¿Desea borrar los datos de la venta en pantalla?", _
              vbQuestion + vbYesNo, "Limpiar venta") = vbYes Then
        LimpiarFormularioVenta
    End If

End Sub

' Borra cliente, productos y cantidades (las formulas se conservan)
Public Sub LimpiarFormularioVenta()

    Dim hoja As Worksheet

    Set hoja = ThisWorkbook.Worksheets(HOJA_VENTA)

    hoja.Range("B3").MergeArea.ClearContents
    hoja.Range(hoja.Cells(FILA_INICIO, COL_PRODUCTO), _
               hoja.Cells(FILA_FIN, COL_CANTIDAD)).ClearContents

End Sub
