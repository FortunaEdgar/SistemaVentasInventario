Attribute VB_Name = "modClientes"
'===============================================================
' Modulo: modClientes
' Registro, edicion y carga de clientes.
'===============================================================
Option Explicit

'---------------------------------------------------------------
' Actualiza la tabla de clientes desde PostgreSQL
'---------------------------------------------------------------
Public Sub CargarClientes()

    On Error GoTo ManejarError

    Application.ScreenUpdating = False
    RefrescarTabla HOJA_CLIENTES, TABLA_CLIENTES, "clientes"
    Application.ScreenUpdating = True
    Exit Sub

ManejarError:
    MostrarError "cargar los clientes", Err.Description

End Sub

'---------------------------------------------------------------
' Registra un cliente nuevo
'---------------------------------------------------------------
Public Sub RegistrarCliente()

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim nombre As String
    Dim telefono As String
    Dim correo As String
    Dim cancelado As Boolean

    Const TITULO As String = "Registrar cliente"

    nombre = PedirTexto("Ingrese el nombre del cliente:", TITULO, cancelado)
    If cancelado Then Exit Sub

    If nombre = "" Then
        MsgBox "El nombre del cliente es obligatorio.", vbExclamation, TITULO
        Exit Sub
    End If

    If Len(nombre) > 100 Then
        MsgBox "El nombre no puede tener mas de 100 caracteres.", vbExclamation, TITULO
        Exit Sub
    End If

    telefono = PedirTexto("Ingrese el telefono del cliente (opcional):", TITULO, cancelado)
    If cancelado Then Exit Sub

    If Len(telefono) > 20 Then
        MsgBox "El telefono no puede tener mas de 20 caracteres.", vbExclamation, TITULO
        Exit Sub
    End If

    correo = PedirTexto("Ingrese el correo del cliente (opcional):", TITULO, cancelado)
    If cancelado Then Exit Sub

    If Not CorreoValido(correo) Then
        MsgBox "El correo no tiene un formato valido (ejemplo: nombre@correo.com).", _
               vbExclamation, TITULO
        Exit Sub
    End If

    On Error GoTo ManejarError

    Set conexion = AbrirConexion()

    Set comando = NuevoComando(conexion, _
        "INSERT INTO clientes (nombre, telefono, correo) VALUES (?, ?, ?);")

    ParamTexto comando, "pNombre", nombre, 100
    ParamTexto comando, "pTelefono", telefono, 20
    ParamTexto comando, "pCorreo", correo, 150

    comando.Execute

    CerrarConexion conexion
    Set comando = Nothing

    CargarClientes
    ActualizarIndicadores

    MsgBox "Cliente """ & nombre & """ registrado correctamente.", vbInformation, TITULO
    Exit Sub

ManejarError:
    MostrarError "registrar el cliente", Err.Description
    CerrarConexion conexion

End Sub

'---------------------------------------------------------------
' Edita el cliente seleccionado en la tabla (o el que se indique)
'---------------------------------------------------------------
Public Sub EditarCliente()

    Dim conexion As ADODB.Connection
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset
    Dim idCliente As Long
    Dim nombre As String
    Dim telefono As String
    Dim correo As String
    Dim cancelado As Boolean

    Const TITULO As String = "Editar cliente"

    On Error GoTo ManejarError

    idCliente = IdFilaSeleccionada( _
        ThisWorkbook.Worksheets(HOJA_CLIENTES).ListObjects(TABLA_CLIENTES))

    If idCliente = 0 Then
        idCliente = PedirEntero("Escriba el ID del cliente que desea editar:" & vbCrLf & _
                                "(Tambien puede seleccionar una fila de la tabla antes de presionar el boton)", _
                                TITULO)
        If idCliente <= 0 Then Exit Sub
    End If

    ' Datos actuales del cliente
    Set conexion = AbrirConexion()
    Set comando = NuevoComando(conexion, _
        "SELECT nombre, telefono, correo FROM clientes WHERE id_cliente = ?;")
    ParamEntero comando, "pId", idCliente
    Set registros = comando.Execute

    If registros.EOF Then
        registros.Close
        CerrarConexion conexion
        MsgBox "No existe un cliente con el ID " & idCliente & ".", vbExclamation, TITULO
        Exit Sub
    End If

    nombre = Nz(registros.Fields("nombre").Value)
    telefono = Nz(registros.Fields("telefono").Value)
    correo = Nz(registros.Fields("correo").Value)
    registros.Close
    Set registros = Nothing

    ' Nuevos datos (se muestran los actuales para modificarlos)
    nombre = PedirTexto("Nombre del cliente:", TITULO & " #" & idCliente, cancelado, nombre)
    If cancelado Then GoTo Salir

    If nombre = "" Or Len(nombre) > 100 Then
        MsgBox "El nombre es obligatorio y no puede pasar de 100 caracteres.", vbExclamation, TITULO
        GoTo Salir
    End If

    telefono = PedirTexto("Telefono (deje vacio para borrarlo):", TITULO & " #" & idCliente, _
                          cancelado, telefono)
    If cancelado Then GoTo Salir

    If Len(telefono) > 20 Then
        MsgBox "El telefono no puede tener mas de 20 caracteres.", vbExclamation, TITULO
        GoTo Salir
    End If

    correo = PedirTexto("Correo (deje vacio para borrarlo):", TITULO & " #" & idCliente, _
                        cancelado, correo)
    If cancelado Then GoTo Salir

    If Not CorreoValido(correo) Then
        MsgBox "El correo no tiene un formato valido.", vbExclamation, TITULO
        GoTo Salir
    End If

    Set comando = NuevoComando(conexion, _
        "UPDATE clientes SET nombre = ?, telefono = ?, correo = ? WHERE id_cliente = ?;")
    ParamTexto comando, "pNombre", nombre, 100
    ParamTexto comando, "pTelefono", telefono, 20
    ParamTexto comando, "pCorreo", correo, 150
    ParamEntero comando, "pId", idCliente
    comando.Execute

    CerrarConexion conexion

    CargarClientes
    MsgBox "Cliente actualizado correctamente.", vbInformation, TITULO

Salir:
    CerrarConexion conexion
    Exit Sub

ManejarError:
    MostrarError "editar el cliente", Err.Description
    CerrarConexion conexion

End Sub
