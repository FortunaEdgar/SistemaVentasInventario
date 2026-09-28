Attribute VB_Name = "modConexion"
'===============================================================
' Modulo: modConexion
' Conexion con PostgreSQL (ODBC + ADO) y funciones de apoyo que
' usan todos los demas modulos.
'
' Requiere la referencia:
'   Microsoft ActiveX Data Objects 6.1 Library
'===============================================================
Option Explicit

' Cadena de conexion unica para todo el sistema.
' Si cambia la clave o el DSN, solo se cambia aqui.
Public Const CADENA_CONEXION As String = _
    "DSN=SistemaVentas_PostgreSQL;UID=postgres;PWD=postgres123;"

' Tasa de ITBIS (debe coincidir con la funcion tasa_itbis() en PostgreSQL).
' Los precios de los productos ya incluyen el ITBIS.
Public Const TASA_ITBIS As Double = 0.18

' Nombres de las hojas del libro
Public Const HOJA_MENU As String = "Menú"
Public Const HOJA_VENTA As String = "Nueva Venta"
Public Const HOJA_FACTURA As String = "Factura"
Public Const HOJA_HISTORIAL As String = "Historial"
Public Const HOJA_REPORTES As String = "Reportes"
Public Const HOJA_CLIENTES As String = "Clientes"
Public Const HOJA_PRODUCTOS As String = "Productos"
Public Const HOJA_VENTAS As String = "Ventas"
Public Const HOJA_DETALLE As String = "Detalle Venta"

' Nombres de las tablas de Excel (cargadas con Power Query)
Public Const TABLA_CLIENTES As String = "clientes"
Public Const TABLA_PRODUCTOS As String = "Productos"
Public Const TABLA_VENTAS As String = "Ventas"
Public Const TABLA_DETALLE As String = "DetalleVentas"

' Formato de moneda usado en mensajes
Public Const FORMATO_MONEDA As String = "RD$ #,##0.00"

'---------------------------------------------------------------
' Abre y devuelve una conexion. Si falla, lanza un error con un
' mensaje claro para el usuario.
'---------------------------------------------------------------
Public Function AbrirConexion() As ADODB.Connection

    Dim conexion As ADODB.Connection
    Dim detalle As String

    Set conexion = New ADODB.Connection
    conexion.ConnectionString = CADENA_CONEXION
    conexion.ConnectionTimeout = 8

    On Error GoTo ErrorConexion
    conexion.Open
    On Error GoTo 0

    Set AbrirConexion = conexion
    Exit Function

ErrorConexion:
    detalle = Err.Description
    Set conexion = Nothing
    On Error GoTo 0
    Err.Raise vbObjectError + 1000, "AbrirConexion", _
        "No se pudo conectar con la base de datos PostgreSQL." & vbCrLf & vbCrLf & _
        "Verifique que:" & vbCrLf & _
        "  1. Docker Desktop este abierto y el contenedor encendido" & vbCrLf & _
        "     (docker compose up -d en la carpeta Docker)." & vbCrLf & _
        "  2. Exista el DSN ODBC 'SistemaVentas_PostgreSQL'." & vbCrLf & vbCrLf & _
        "Detalle tecnico: " & detalle

End Function

'---------------------------------------------------------------
' Cierra una conexion sin generar errores si ya estaba cerrada.
'---------------------------------------------------------------
Public Sub CerrarConexion(ByRef conexion As ADODB.Connection)

    On Error Resume Next

    If Not conexion Is Nothing Then
        If conexion.State = adStateOpen Then conexion.Close
    End If

    Set conexion = Nothing

End Sub

'---------------------------------------------------------------
' Crea un comando parametrizado listo para agregar parametros.
'---------------------------------------------------------------
Public Function NuevoComando(ByVal conexion As ADODB.Connection, _
                             ByVal sql As String) As ADODB.Command

    Dim comando As ADODB.Command

    Set comando = New ADODB.Command
    Set comando.ActiveConnection = conexion
    comando.CommandText = sql
    comando.CommandType = adCmdText

    Set NuevoComando = comando

End Function

' Agrega un parametro de texto (un texto vacio se guarda como NULL)
Public Sub ParamTexto(ByVal comando As ADODB.Command, ByVal nombre As String, _
                      ByVal valor As String, ByVal tamano As Long)

    If Trim$(valor) = "" Then
        comando.Parameters.Append comando.CreateParameter( _
            nombre, adVarChar, adParamInput, tamano, Null)
    Else
        comando.Parameters.Append comando.CreateParameter( _
            nombre, adVarChar, adParamInput, tamano, Trim$(valor))
    End If

End Sub

' Agrega un parametro entero
Public Sub ParamEntero(ByVal comando As ADODB.Command, ByVal nombre As String, _
                       ByVal valor As Long)

    comando.Parameters.Append comando.CreateParameter( _
        nombre, adInteger, adParamInput, , valor)

End Sub

' Agrega un parametro decimal
Public Sub ParamDecimal(ByVal comando As ADODB.Command, ByVal nombre As String, _
                        ByVal valor As Double)

    comando.Parameters.Append comando.CreateParameter( _
        nombre, adDouble, adParamInput, , valor)

End Sub

'---------------------------------------------------------------
' Convierte el error que devuelve PostgreSQL/ODBC en un mensaje
' entendible para el usuario.
'---------------------------------------------------------------
Public Function MensajeBD(ByVal descripcion As String) As String

    Dim texto As String
    Dim posicion As Long

    ' Error de conexion (ya viene con un mensaje claro)
    If Left$(descripcion, 19) = "No se pudo conectar" Then
        MensajeBD = descripcion
        Exit Function
    End If

    ' Errores de restricciones conocidas
    If InStr(1, descripcion, "uq_clientes_nombre", vbTextCompare) > 0 Then
        MensajeBD = "Ya existe un cliente con ese nombre."
        Exit Function
    End If

    If InStr(1, descripcion, "uq_productos_nombre", vbTextCompare) > 0 Then
        MensajeBD = "Ya existe un producto con ese nombre."
        Exit Function
    End If

    If InStr(1, descripcion, "ck_clientes_correo", vbTextCompare) > 0 Then
        MensajeBD = "El correo no tiene un formato valido."
        Exit Function
    End If

    If InStr(1, descripcion, "ck_productos_precio", vbTextCompare) > 0 Then
        MensajeBD = "El precio debe ser mayor que cero."
        Exit Function
    End If

    If InStr(1, descripcion, "ck_productos_stock", vbTextCompare) > 0 Then
        MensajeBD = "El stock no puede quedar negativo."
        Exit Function
    End If

    If InStr(1, descripcion, "fk_ventas_cliente", vbTextCompare) > 0 Then
        MensajeBD = "El cliente tiene ventas registradas."
        Exit Function
    End If

    ' Mensajes lanzados con RAISE EXCEPTION en las funciones:
    ' "ERROR: Stock insuficiente ...;" -> "Stock insuficiente ..."
    texto = descripcion
    posicion = InStr(1, texto, "ERROR:", vbTextCompare)

    If posicion > 0 Then
        texto = Mid$(texto, posicion + 6)

        posicion = InStr(texto, ";")
        If posicion > 0 Then texto = Left$(texto, posicion - 1)

        posicion = InStr(texto, vbLf)
        If posicion > 0 Then texto = Left$(texto, posicion - 1)
    End If

    MensajeBD = Trim$(texto)

End Function

' Muestra un error con formato uniforme
Public Sub MostrarError(ByVal accion As String, ByVal descripcion As String)

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

    MsgBox "No se pudo " & accion & "." & vbCrLf & vbCrLf & _
           MensajeBD(descripcion), vbCritical, "Sistema de Ventas"

End Sub

'---------------------------------------------------------------
' Pide un texto al usuario.
' Devuelve False en "cancelado" si el usuario presiona Cancelar.
'---------------------------------------------------------------
Public Function PedirTexto(ByVal mensaje As String, ByVal titulo As String, _
                           ByRef cancelado As Boolean, _
                           Optional ByVal valorInicial As String = "") As String

    Dim respuesta As Variant

    respuesta = Application.InputBox(mensaje, titulo, valorInicial, Type:=2)

    If VarType(respuesta) = vbBoolean Then
        cancelado = True
        PedirTexto = ""
    Else
        cancelado = False
        PedirTexto = Trim$(CStr(respuesta))
    End If

End Function

' Pide un numero entero. Devuelve -1 si se cancela o no es valido.
Public Function PedirEntero(ByVal mensaje As String, ByVal titulo As String, _
                            Optional ByVal valorInicial As String = "") As Long

    Dim texto As String
    Dim cancelado As Boolean

    texto = PedirTexto(mensaje, titulo, cancelado, valorInicial)

    If cancelado Or texto = "" Then
        PedirEntero = -1
    ElseIf Not EsEntero(texto) Then
        MsgBox "Debe escribir un numero entero valido.", vbExclamation, titulo
        PedirEntero = -1
    Else
        PedirEntero = CLng(texto)
    End If

End Function

' Verifica que un valor sea un numero entero (sin decimales)
Public Function EsEntero(ByVal valor As Variant) As Boolean

    EsEntero = False

    If IsError(valor) Or IsEmpty(valor) Then Exit Function
    If Not IsNumeric(valor) Then Exit Function
    If CDbl(valor) <> Fix(CDbl(valor)) Then Exit Function
    If Abs(CDbl(valor)) > 2147483647# Then Exit Function

    EsEntero = True

End Function

' Devuelve el texto de una celda (vacio si la celda tiene error)
Public Function TextoCelda(ByVal celda As Range) As String

    If IsError(celda.Value) Then
        TextoCelda = ""
    Else
        TextoCelda = Trim$(CStr(celda.Value))
    End If

End Function

' Convierte un valor de la base de datos (que puede ser NULL) a texto
Public Function Nz(ByVal valor As Variant) As String

    If IsNull(valor) Or IsEmpty(valor) Then
        Nz = ""
    Else
        Nz = CStr(valor)
    End If

End Function

' Validacion sencilla de correo: algo@algo.algo
Public Function CorreoValido(ByVal correo As String) As Boolean

    If correo = "" Then
        CorreoValido = True
    Else
        CorreoValido = (correo Like "*?@?*.?*") And (InStr(correo, " ") = 0)
    End If

End Function

' Numero de factura con formato FAC-000001
Public Function NumeroFactura(ByVal idVenta As Long) As String

    NumeroFactura = "FAC-" & Format$(idVenta, "000000")

End Function

' Convierte "FAC-000005", "5" o "000005" en 5. Devuelve 0 si no es valido.
Public Function IdDesdeNumeroFactura(ByVal texto As String) As Long

    Dim i As Long
    Dim digitos As String

    For i = 1 To Len(texto)
        If Mid$(texto, i, 1) Like "#" Then digitos = digitos & Mid$(texto, i, 1)
    Next i

    If digitos = "" Or Len(digitos) > 9 Then
        IdDesdeNumeroFactura = 0
    Else
        IdDesdeNumeroFactura = CLng(digitos)
    End If

End Function

'---------------------------------------------------------------
' Devuelve el ID (primera columna) de la fila seleccionada dentro
' de una tabla de Excel. Devuelve 0 si no hay fila seleccionada.
'---------------------------------------------------------------
Public Function IdFilaSeleccionada(ByVal tabla As ListObject) As Long

    Dim fila As Long
    Dim valor As Variant

    IdFilaSeleccionada = 0

    If ActiveSheet.Name <> tabla.Parent.Name Then Exit Function
    If tabla.DataBodyRange Is Nothing Then Exit Function
    If Intersect(ActiveCell, tabla.DataBodyRange) Is Nothing Then Exit Function

    fila = ActiveCell.Row - tabla.DataBodyRange.Row + 1
    valor = tabla.DataBodyRange.Cells(fila, 1).Value

    If EsEntero(valor) Then IdFilaSeleccionada = CLng(valor)

End Function

'---------------------------------------------------------------
' Actualiza una tabla de Excel con los datos de PostgreSQL.
' Primero intenta con Power Query (la conexion del libro); si no
' esta disponible, carga los datos directamente con ADO.
'---------------------------------------------------------------
Public Sub RefrescarTabla(ByVal nombreHoja As String, ByVal nombreTabla As String, _
                          ByVal tablaBD As String)

    Dim tabla As ListObject
    Dim correcto As Boolean

    Set tabla = ThisWorkbook.Worksheets(nombreHoja).ListObjects(nombreTabla)

    On Error Resume Next
    tabla.QueryTable.Refresh BackgroundQuery:=False
    correcto = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    If Not correcto Then CargarTablaConADO tabla, tablaBD

End Sub

' Carga una tabla usando los encabezados como nombres de columna
Private Sub CargarTablaConADO(ByVal tabla As ListObject, ByVal tablaBD As String)

    Dim conexion As ADODB.Connection
    Dim registros As ADODB.Recordset
    Dim columnas As String
    Dim i As Long
    Dim filas As Long

    For i = 1 To tabla.ListColumns.Count
        If i > 1 Then columnas = columnas & ", "
        columnas = columnas & tabla.ListColumns(i).Name
    Next i

    Set conexion = AbrirConexion()
    Set registros = New ADODB.Recordset
    registros.CursorLocation = adUseClient
    registros.Open "SELECT " & columnas & " FROM " & tablaBD & _
                   " ORDER BY 1;", conexion, adOpenStatic, adLockReadOnly

    filas = registros.RecordCount

    If Not tabla.DataBodyRange Is Nothing Then tabla.DataBodyRange.ClearContents

    If filas > 0 Then
        tabla.HeaderRowRange.Cells(1, 1).Offset(1, 0).CopyFromRecordset registros
        On Error Resume Next
        tabla.Resize tabla.HeaderRowRange.Resize(filas + 1)
        On Error GoTo 0
    End If

    registros.Close
    Set registros = Nothing
    CerrarConexion conexion

End Sub
