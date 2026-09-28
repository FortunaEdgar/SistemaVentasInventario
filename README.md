# Sistema de Ventas e Inventario

Sistema de facturación e inventario hecho con **PostgreSQL** (en Docker) y **Excel + VBA**.
Excel es la interfaz: se conecta a la base de datos por ODBC, y la lógica importante
(validar stock, descontar el inventario, calcular el ITBIS y anular ventas) se ejecuta dentro de
PostgreSQL en una sola transacción.

## Funcionalidades

| Módulo | Qué hace |
|---|---|
| **Menú** | Navegación, indicadores (clientes, productos, stock bajo, ventas del día y del mes) y datos del negocio que se imprimen en la factura. |
| **Clientes** | Registrar, editar y listar clientes (nombre único, validación de correo). |
| **Productos** | Registrar, editar, **entradas de inventario** y listado con alerta de stock bajo (≤ 5). |
| **Nueva Venta** | Venta con hasta 21 productos, precio y stock automáticos, total con **ITBIS 18 % incluido** en el precio y su desglose. |
| **Factura** | Factura imprimible (FAC-000001), búsqueda, **exportación a PDF** y anulación. |
| **Historial** | Listado de ventas con filtros por fecha y cliente; ver o anular una factura. |
| **Reportes** | Resumen por período, ventas por día, productos más vendidos, stock bajo y kardex. |

## Estructura del proyecto

```
Docker/docker-compose.yml        Contenedor de PostgreSQL 16
SQL/01_crear_tablas.sql          Tablas, llaves y restricciones
SQL/02_funciones_y_vistas.sql    Funciones (registrar_venta, anular_venta, entrada_inventario),
                                 trigger del kardex, vistas y funciones de reportes
SQL/03_datos_prueba.sql          Datos de ejemplo
Excel/SistemaVentasInventario.xlsm   Libro con las hojas, botones y macros
VBA/*.bas, VBA/ThisWorkbook.cls  Copia del código VBA (para ver los cambios en Git)
```

### Base de datos

- `clientes`, `productos`, `ventas`, `detalle_venta` y `movimientos_inventario` (kardex).
- Restricciones: precio > 0, stock ≥ 0, cantidad > 0, nombres únicos, formato de correo,
  `total = subtotal + itbis`, estado `EMITIDA` / `ANULADA`.
- `registrar_venta(cliente, productos[], cantidades[])`: guarda el encabezado y el detalle,
  bloquea los productos (`FOR UPDATE`), valida y descuenta el stock, registra el kardex y desglosa el ITBIS.
  Los precios de los productos **ya incluyen el ITBIS**: una venta de 1,180.00 se registra como
  subtotal 1,000.00 + ITBIS 180.00.
  Si una línea falla, no se guarda nada.
- `anular_venta(id)`: marca la venta como anulada y devuelve los productos al inventario.
- `entrada_inventario(id, cantidad, referencia)`: suma stock y lo registra en el kardex.

### Código VBA

| Módulo | Contenido |
|---|---|
| `modConexion` | Cadena de conexión única, apertura y cierre, parámetros ADO, mensajes de error claros y utilidades. |
| `modMenu` | Inicio del sistema, indicadores, "Actualizar datos", "Probar conexión" y navegación. |
| `modClientes` | `RegistrarCliente`, `EditarCliente`, `CargarClientes`. |
| `modProductos` | `RegistrarProducto`, `EditarProducto`, `EntradaInventario`, `CargarProductos`. |
| `modVentas` | `RegistrarVenta`, `LimpiarVenta`. |
| `modFactura` | `MostrarFactura`, `BuscarFactura`, `ExportarFacturaPDF`, `AnularFacturaActual`. |
| `modHistorial` | `CargarHistorial`, `VerFacturaSeleccionada`, `AnularVentaSeleccionada`. |
| `modReportes` | `GenerarReportes`. |

Todas las consultas usan **parámetros ADO** (`?`), así se evita la inyección SQL.

## Instalación

### 1. Base de datos (Docker)

```bash
cd Docker
docker compose down -v      # borra el volumen anterior (¡elimina los datos de prueba viejos!)
docker compose up -d
```

La primera vez que se crea el volumen, Docker ejecuta automáticamente los scripts de la carpeta
`SQL` en orden (01, 02, 03). La base se llama `SistemaVentas` (usuario `postgres`, clave `postgres123`)
y usa la zona horaria `America/Santo_Domingo`.

Para reinstalar sin borrar el volumen se pueden ejecutar los tres scripts manualmente en
pgAdmin o DBeaver, en orden, sobre la base `SistemaVentas`.

### 2. Conexión ODBC

1. Instalar el driver **psqlODBC** (64 bits si Office es de 64 bits).
2. Abrir *Orígenes de datos ODBC* → *DSN de usuario* → *Agregar* → **PostgreSQL Unicode(x64)**.
3. Datos: Nombre `SistemaVentas_PostgreSQL`, Servidor `localhost`, Puerto `5432`,
   Base de datos `SistemaVentas`, Usuario `postgres`, Clave `postgres123`.

### 3. Excel

1. Abrir `Excel/SistemaVentasInventario.xlsm` y **habilitar el contenido (macros)**.
   Si Windows bloquea las macros por venir de Internet: clic derecho en el archivo →
   *Propiedades* → marcar *Desbloquear*.
2. Verificar en el editor de VBA (*Herramientas → Referencias*) que esté marcada
   **Microsoft ActiveX Data Objects 6.1 Library**.
3. Al abrir, el libro muestra el **Menú** y actualiza los datos. Use *Probar conexión* si algo falla.

## Uso rápido

1. **Productos → Registrar producto** (o *Entrada de inventario* para sumar stock).
2. **Clientes → Registrar cliente**.
3. **Nueva Venta**: elegir el cliente, los productos y las cantidades → **Registrar venta**.
4. Se abre la **Factura**; se puede guardar en PDF (carpeta `Excel/Facturas`).
5. **Historial** para consultar o anular ventas y **Reportes** para ver resultados del período.

## Autor

Edgar Antonio Yan Fortuna — Universidad Federico Henríquez y Carvajal (UFHEC), Campus La Romana.
