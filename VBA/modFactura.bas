Attribute VB_Name = "modFactura"
'===============================================================
' Modulo: modFactura
' Muestra, busca y exporta a PDF la factura de una venta.
'
' Distribucion de la hoja "Factura":
'   E2 numero, E3 fecha, E4 estado
'   B7 cliente, B8 telefono, B9 correo
'   A12:E32 lineas (codigo, descripcion, cantidad, precio, importe)
'   E34 subtotal, E35 ITBIS, E36 total
'===============================================================
Option Explicit

Private Const FILA_LINEAS As Long = 12
Private Const MAX_LINEAS As Long = 21

'---------------------------------------------------------------
' Carga en la hoja "Factura" la venta indicada
'---------------------------------------------------------------
Public Sub MostrarFactura(ByVal idVenta As Long)

    Dim hoja As Worksheet
    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset

    On Error GoTo ManejarError

    Set hoja = ThisWorkbook.Worksheets(HOJA_FACTURA)
    Set conexion = AbrirConexion()

    ' Encabezado
    Set comando = NuevoComando(conexion, _
        "SELECT numero, fecha, estado, cliente, telefono, correo, " & _
        "subtotal, itbis, total FROM v_factura_encabezado WHERE id_venta = ?;")
    ParamEntero comando, "pId", idVenta
    Set registros = comando.Execute

    If registros.EOF Then
        registros.Close
        CerrarConexion conexion
        MsgBox "No existe la factura " & NumeroFactura(idVenta) & ".", _
               vbExclamation, "Factura"
        Exit Sub
    End If

    LimpiarFactura

    hoja.Range("E2").Value = registros.Fields("numero").Value
    hoja.Range("E3").Value = registros.Fields("fecha").Value
    hoja.Range("E4").Value = registros.Fields("estado").Value
    hoja.Range("B7").Value = Nz(registros.Fields("cliente").Value)
    hoja.Range("B8").Value = Nz(registros.Fields("telefono").Value)
    hoja.Range("B9").Value = Nz(registros.Fields("correo").Value)
    hoja.Range("E34").Value = registros.Fields("subtotal").Value
    hoja.Range("E35").Value = registros.Fields("itbis").Value
    hoja.Range("E36").Value = registros.Fields("total").Value

    registros.Close

    ' Lineas
    Set comando = NuevoComando(conexion, _
        "SELECT id_producto, producto, cantidad, precio, subtotal " & _
        "FROM v_factura_detalle WHERE id_venta = ? ORDER BY id_detalle;")
    ParamEntero comando, "pId", idVenta
    Set registros = comando.Execute

    hoja.Cells(FILA_LINEAS, 1).CopyFromRecordset registros, MAX_LINEAS

    registros.Close
    Set registros = Nothing
    CerrarConexion conexion

    hoja.Activate
    hoja.Range("A1").Select
    Exit Sub

ManejarError:
    MostrarError "mostrar la factura", Err.Description
    CerrarConexion conexion

End Sub

' Deja la factura en blanco
Public Sub LimpiarFactura()

    With ThisWorkbook.Worksheets(HOJA_FACTURA)
        .Range("E2:E4").ClearContents
        .Range("B7:C9").ClearContents
        .Range(.Cells(FILA_LINEAS, 1), .Cells(FILA_LINEAS + MAX_LINEAS - 1, 5)).ClearContents
        .Range("E34:E36").ClearContents
    End With

End Sub

'---------------------------------------------------------------
' Boton "Buscar factura"
'---------------------------------------------------------------
Public Sub BuscarFactura()

    Dim texto As String
    Dim idVenta As Long
    Dim cancelado As Boolean

    texto = PedirTexto("Escriba el numero de la factura" & vbCrLf & _
                       "(ejemplo: 5 o FAC-000005):", "Buscar factura", cancelado)
    If cancelado Or texto = "" Then Exit Sub

    idVenta = IdDesdeNumeroFactura(texto)

    If idVenta <= 0 Then
        MsgBox "Numero de factura no valido.", vbExclamation, "Buscar factura"
        Exit Sub
    End If

    MostrarFactura idVenta

End Sub

' Id de la factura que esta en pantalla (0 si no hay ninguna)
Public Function IdFacturaActual() As Long

    IdFacturaActual = IdDesdeNumeroFactura( _
        TextoCelda(ThisWorkbook.Worksheets(HOJA_FACTURA).Range("E2")))

End Function

'---------------------------------------------------------------
' Boton "Exportar PDF": guarda la factura en la carpeta
' "Facturas" junto al archivo de Excel.
'---------------------------------------------------------------
Public Sub ExportarFacturaPDF()

    Dim hoja As Worksheet
    Dim carpeta As String
    Dim archivo As String
    Dim numero As String

    On Error GoTo ManejarError

    Set hoja = ThisWorkbook.Worksheets(HOJA_FACTURA)
    numero = TextoCelda(hoja.Range("E2"))

    If numero = "" Then
        MsgBox "Primero busque o registre una factura.", vbExclamation, "Exportar PDF"
        Exit Sub
    End If

    carpeta = ThisWorkbook.Path

    ' Si el libro esta en OneDrive (ruta web) o sin guardar, usar Documentos
    If carpeta = "" Or LCase$(Left$(carpeta, 4)) = "http" Then
        carpeta = Environ$("USERPROFILE") & "\Documents"
    End If

    carpeta = carpeta & "\Facturas"
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta

    archivo = carpeta & "\" & numero & ".pdf"

    hoja.ExportAsFixedFormat Type:=xlTypePDF, Filename:=archivo, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True

    Exit Sub

ManejarError:
    MostrarError "exportar la factura a PDF", Err.Description

End Sub

'---------------------------------------------------------------
' Boton "Anular factura" (anula la factura que esta en pantalla)
'---------------------------------------------------------------
Public Sub AnularFacturaActual()

    Dim idVenta As Long

    idVenta = IdFacturaActual()

    If idVenta = 0 Then
        MsgBox "Primero busque la factura que desea anular.", vbExclamation, "Anular factura"
        Exit Sub
    End If

    If AnularVentaPorId(idVenta) Then MostrarFactura idVenta

End Sub
