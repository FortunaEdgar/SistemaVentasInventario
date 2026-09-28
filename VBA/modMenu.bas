Attribute VB_Name = "modMenu"
'===============================================================
' Modulo: modMenu
' Menu principal: navegacion, indicadores y actualizacion general.
'
' Indicadores en la hoja "Menú":
'   B17 clientes, D17 productos, F17 stock bajo
'   B20 ventas de hoy, D20 total de hoy, F20 total del mes
'   B22 fecha de la ultima actualizacion
'===============================================================
Option Explicit

'---------------------------------------------------------------
' Se ejecuta al abrir el libro (desde ThisWorkbook)
'---------------------------------------------------------------
Public Sub IniciarSistema()

    Dim conexion As ADODB.Connection

    ThisWorkbook.Worksheets(HOJA_MENU).Activate

    ' Si la base de datos no esta disponible, se avisa una sola vez
    On Error GoTo SinConexion
    Set conexion = AbrirConexion()
    CerrarConexion conexion
    On Error GoTo 0

    ActualizarDatos False
    Exit Sub

SinConexion:
    MsgBox "El libro se abrio sin conexion a la base de datos." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Sistema de Ventas"
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Boton "Actualizar datos"
'---------------------------------------------------------------
Public Sub ActualizarTodo()

    ActualizarDatos True

End Sub

' Refresca todas las tablas y los indicadores
Public Sub ActualizarDatos(ByVal avisar As Boolean)

    On Error GoTo ManejarError

    Application.ScreenUpdating = False
    Application.Cursor = xlWait

    RefrescarTabla HOJA_CLIENTES, TABLA_CLIENTES, "clientes"
    RefrescarTabla HOJA_PRODUCTOS, TABLA_PRODUCTOS, "productos"
    RefrescarTabla HOJA_VENTAS, TABLA_VENTAS, "ventas"
    RefrescarTabla HOJA_DETALLE, TABLA_DETALLE, "detalle_venta"
    ActualizarIndicadores

    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    If avisar Then MsgBox "Datos actualizados correctamente.", vbInformation, "Sistema de Ventas"
    Exit Sub

ManejarError:
    MostrarError "actualizar los datos", Err.Description

End Sub

'---------------------------------------------------------------
' Carga los indicadores del menu (no muestra errores)
'---------------------------------------------------------------
Public Sub ActualizarIndicadores()

    Dim conexion As ADODB.Connection
    Dim registros As ADODB.Recordset

    On Error GoTo Salir

    Set conexion = AbrirConexion()
    Set registros = conexion.Execute("SELECT * FROM v_indicadores;")

    With ThisWorkbook.Worksheets(HOJA_MENU)
        .Range("B17").Value = registros.Fields("total_clientes").Value
        .Range("D17").Value = registros.Fields("total_productos").Value
        .Range("F17").Value = registros.Fields("productos_stock_bajo").Value
        .Range("B20").Value = registros.Fields("ventas_hoy").Value
        .Range("D20").Value = registros.Fields("total_hoy").Value
        .Range("F20").Value = registros.Fields("total_mes").Value
        .Range("B22").Value = "Última actualización: " & Format$(Now, "dd/mm/yyyy hh:mm")
    End With

    registros.Close

Salir:
    Set registros = Nothing
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Boton "Probar conexion"
'---------------------------------------------------------------
Public Sub ProbarConexion()

    Dim conexion As ADODB.Connection
    Dim registros As ADODB.Recordset
    Dim mensaje As String

    On Error GoTo ManejarError

    Set conexion = AbrirConexion()
    Set registros = conexion.Execute( _
        "SELECT current_database() AS base, version() AS version, " & _
        "to_char(now(), 'DD/MM/YYYY HH24:MI') AS hora;")

    mensaje = "Conexion exitosa con PostgreSQL." & vbCrLf & vbCrLf & _
              "Base de datos: " & registros.Fields("base").Value & vbCrLf & _
              "Hora del servidor: " & registros.Fields("hora").Value & vbCrLf & vbCrLf & _
              Left$(Nz(registros.Fields("version").Value), 60)

    registros.Close
    Set registros = Nothing
    CerrarConexion conexion

    MsgBox mensaje, vbInformation, "Probar conexion"
    Exit Sub

ManejarError:
    MostrarError "conectar con la base de datos", Err.Description
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Navegacion entre hojas (botones)
'---------------------------------------------------------------
Public Sub IrAMenu()
    ThisWorkbook.Worksheets(HOJA_MENU).Activate
    ThisWorkbook.Worksheets(HOJA_MENU).Range("A1").Select
End Sub

Public Sub IrANuevaVenta()
    ThisWorkbook.Worksheets(HOJA_VENTA).Activate
    ThisWorkbook.Worksheets(HOJA_VENTA).Range("B3").Select
End Sub

Public Sub IrAFactura()
    ThisWorkbook.Worksheets(HOJA_FACTURA).Activate
End Sub

Public Sub IrAClientes()
    ThisWorkbook.Worksheets(HOJA_CLIENTES).Activate
End Sub

Public Sub IrAProductos()
    ThisWorkbook.Worksheets(HOJA_PRODUCTOS).Activate
End Sub

Public Sub IrAHistorial()
    ThisWorkbook.Worksheets(HOJA_HISTORIAL).Activate
    CargarHistorial
End Sub

Public Sub IrAReportes()
    ThisWorkbook.Worksheets(HOJA_REPORTES).Activate
    GenerarReportes
End Sub
