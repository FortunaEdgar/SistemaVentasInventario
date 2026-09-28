Attribute VB_Name = "modProductos"
'===============================================================
' Modulo: modProductos
' Registro, edicion y entradas de inventario de productos.
'===============================================================
Option Explicit

'---------------------------------------------------------------
' Actualiza la tabla de productos desde PostgreSQL
'---------------------------------------------------------------
Public Sub CargarProductos()

    On Error GoTo ManejarError

    Application.ScreenUpdating = False
    RefrescarTabla HOJA_PRODUCTOS, TABLA_PRODUCTOS, "productos"
    Application.ScreenUpdating = True
    Exit Sub

ManejarError:
    MostrarError "cargar los productos", Err.Description

End Sub

'---------------------------------------------------------------
' Pide y valida un precio. Devuelve -1 si se cancela o no es valido.
'---------------------------------------------------------------
Private Function PedirPrecio(ByVal titulo As String, _
                             Optional ByVal valorInicial As String = "") As Double

    Dim texto As String
    Dim cancelado As Boolean

    PedirPrecio = -1

    texto = PedirTexto("Ingrese el precio de venta del producto (ITBIS incluido):", titulo, _
                       cancelado, valorInicial)
    If cancelado Then Exit Function

    If texto = "" Or Not IsNumeric(texto) Then
        MsgBox "Debe ingresar un precio valido.", vbExclamation, titulo
        Exit Function
    End If

    If CDbl(texto) <= 0 Then
        MsgBox "El precio debe ser mayor que 0.", vbExclamation, titulo
        Exit Function
    End If

    If CDbl(texto) > 99999999.99 Then
        MsgBox "El precio es demasiado alto.", vbExclamation, titulo
        Exit Function
    End If

    PedirPrecio = Round(CDbl(texto), 2)

End Function

'---------------------------------------------------------------
' Registra un producto nuevo
'---------------------------------------------------------------
Public Sub RegistrarProducto()

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim nombre As String
    Dim categoria As String
    Dim stockTexto As String
    Dim precio As Double
    Dim stock As Long
    Dim cancelado As Boolean

    Const TITULO As String = "Registrar producto"

    nombre = PedirTexto("Ingrese el nombre del producto:", TITULO, cancelado)
    If cancelado Then Exit Sub

    If nombre = "" Or Len(nombre) > 150 Then
        MsgBox "El nombre del producto es obligatorio (maximo 150 caracteres).", _
               vbExclamation, TITULO
        Exit Sub
    End If

    categoria = PedirTexto("Ingrese la categoria del producto:", TITULO, cancelado)
    If cancelado Then Exit Sub

    If categoria = "" Or Len(categoria) > 100 Then
        MsgBox "La categoria del producto es obligatoria (maximo 100 caracteres).", _
               vbExclamation, TITULO
        Exit Sub
    End If

    precio = PedirPrecio(TITULO)
    If precio < 0 Then Exit Sub

    stockTexto = PedirTexto("Ingrese el stock inicial del producto:", TITULO, cancelado, "0")
    If cancelado Then Exit Sub

    If Not EsEntero(stockTexto) Then
        MsgBox "Debe ingresar un stock valido (numero entero).", vbExclamation, TITULO
        Exit Sub
    End If

    stock = CLng(stockTexto)

    If stock < 0 Then
        MsgBox "El stock no puede ser negativo.", vbExclamation, TITULO
        Exit Sub
    End If

    On Error GoTo ManejarError

    Set conexion = AbrirConexion()

    Set comando = NuevoComando(conexion, _
        "INSERT INTO productos (nombre, categoria, precio, stock) VALUES (?, ?, ?, ?);")

    ParamTexto comando, "pNombre", nombre, 150
    ParamTexto comando, "pCategoria", categoria, 100
    ParamDecimal comando, "pPrecio", precio
    ParamEntero comando, "pStock", stock

    comando.Execute

    CerrarConexion conexion
    Set comando = Nothing

    CargarProductos
    ActualizarIndicadores

    MsgBox "Producto """ & nombre & """ registrado correctamente.", vbInformation, TITULO
    Exit Sub

ManejarError:
    MostrarError "registrar el producto", Err.Description
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Obtiene el ID del producto seleccionado o lo pide al usuario
'---------------------------------------------------------------
Private Function ProductoSeleccionado(ByVal titulo As String) As Long

    Dim idProducto As Long

    idProducto = IdFilaSeleccionada( _
        ThisWorkbook.Worksheets(HOJA_PRODUCTOS).ListObjects(TABLA_PRODUCTOS))

    If idProducto = 0 Then
        idProducto = PedirEntero("Escriba el ID del producto:" & vbCrLf & _
                                 "(Tambien puede seleccionar una fila de la tabla antes de presionar el boton)", _
                                 titulo)
    End If

    ProductoSeleccionado = idProducto

End Function

'---------------------------------------------------------------
' Edita nombre, categoria y precio de un producto.
' El stock NO se edita aqui: se modifica con ventas y entradas
' de inventario para que quede registrado en el kardex.
'---------------------------------------------------------------
Public Sub EditarProducto()

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset
    Dim idProducto As Long
    Dim nombre As String
    Dim categoria As String
    Dim precio As Double
    Dim cancelado As Boolean

    Const TITULO As String = "Editar producto"

    On Error GoTo ManejarError

    idProducto = ProductoSeleccionado(TITULO)
    If idProducto <= 0 Then Exit Sub

    Set conexion = AbrirConexion()
    Set comando = NuevoComando(conexion, _
        "SELECT nombre, categoria, precio FROM productos WHERE id_producto = ?;")
    ParamEntero comando, "pId", idProducto
    Set registros = comando.Execute

    If registros.EOF Then
        registros.Close
        CerrarConexion conexion
        MsgBox "No existe un producto con el ID " & idProducto & ".", vbExclamation, TITULO
        Exit Sub
    End If

    nombre = Nz(registros.Fields("nombre").Value)
    categoria = Nz(registros.Fields("categoria").Value)
    precio = CDbl(registros.Fields("precio").Value)
    registros.Close
    Set registros = Nothing

    nombre = PedirTexto("Nombre del producto:", TITULO & " #" & idProducto, cancelado, nombre)
    If cancelado Then GoTo Salir

    If nombre = "" Or Len(nombre) > 150 Then
        MsgBox "El nombre es obligatorio (maximo 150 caracteres).", vbExclamation, TITULO
        GoTo Salir
    End If

    categoria = PedirTexto("Categoria:", TITULO & " #" & idProducto, cancelado, categoria)
    If cancelado Then GoTo Salir

    If categoria = "" Or Len(categoria) > 100 Then
        MsgBox "La categoria es obligatoria (maximo 100 caracteres).", vbExclamation, TITULO
        GoTo Salir
    End If

    precio = PedirPrecio(TITULO & " #" & idProducto, CStr(precio))
    If precio < 0 Then GoTo Salir

    Set comando = NuevoComando(conexion, _
        "UPDATE productos SET nombre = ?, categoria = ?, precio = ? WHERE id_producto = ?;")
    ParamTexto comando, "pNombre", nombre, 150
    ParamTexto comando, "pCategoria", categoria, 100
    ParamDecimal comando, "pPrecio", precio
    ParamEntero comando, "pId", idProducto
    comando.Execute

    CerrarConexion conexion

    CargarProductos
    MsgBox "Producto actualizado correctamente." & vbCrLf & _
           "(Las ventas anteriores conservan el precio con que se facturaron.)", _
           vbInformation, TITULO

Salir:
    CerrarConexion conexion
    Exit Sub

ManejarError:
    MostrarError "editar el producto", Err.Description
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Entrada de mercancia: suma unidades al stock de un producto
'---------------------------------------------------------------
Public Sub EntradaInventario()

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset
    Dim idProducto As Long
    Dim cantidad As Long
    Dim nombre As String
    Dim stockActual As Long
    Dim referencia As String
    Dim cancelado As Boolean

    Const TITULO As String = "Entrada de inventario"

    On Error GoTo ManejarError

    idProducto = ProductoSeleccionado(TITULO)
    If idProducto <= 0 Then Exit Sub

    Set conexion = AbrirConexion()
    Set comando = NuevoComando(conexion, _
        "SELECT nombre, stock FROM productos WHERE id_producto = ?;")
    ParamEntero comando, "pId", idProducto
    Set registros = comando.Execute

    If registros.EOF Then
        registros.Close
        CerrarConexion conexion
        MsgBox "No existe un producto con el ID " & idProducto & ".", vbExclamation, TITULO
        Exit Sub
    End If

    nombre = Nz(registros.Fields("nombre").Value)
    stockActual = CLng(registros.Fields("stock").Value)
    registros.Close
    Set registros = Nothing

    cantidad = PedirEntero("Producto: " & nombre & vbCrLf & _
                           "Stock actual: " & stockActual & vbCrLf & vbCrLf & _
                           "Cantidad que entra al inventario:", TITULO)
    If cantidad < 0 Then GoTo Salir

    If cantidad = 0 Then
        MsgBox "La cantidad debe ser mayor que cero.", vbExclamation, TITULO
        GoTo Salir
    End If

    referencia = PedirTexto("Referencia (opcional, ej. factura del proveedor):", TITULO, _
                            cancelado, "Compra a proveedor")
    If cancelado Then GoTo Salir

    Set comando = NuevoComando(conexion, _
        "SELECT entrada_inventario(?, ?, CAST(? AS VARCHAR)) AS nuevo_stock;")
    ParamEntero comando, "pId", idProducto
    ParamEntero comando, "pCantidad", cantidad
    ParamTexto comando, "pReferencia", Left$(referencia, 150), 150
    Set registros = comando.Execute

    stockActual = CLng(registros.Fields("nuevo_stock").Value)
    registros.Close
    Set registros = Nothing

    CerrarConexion conexion

    CargarProductos
    ActualizarIndicadores

    MsgBox "Entrada registrada." & vbCrLf & vbCrLf & _
           nombre & vbCrLf & "Nuevo stock: " & stockActual, vbInformation, TITULO

Salir:
    CerrarConexion conexion
    Exit Sub

ManejarError:
    MostrarError "registrar la entrada de inventario", Err.Description
    CerrarConexion conexion

End Sub
