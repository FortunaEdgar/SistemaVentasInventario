-- =============================================================
--  Sistema de Ventas e Inventario
--  01_crear_tablas.sql
--  Crea la estructura de la base de datos SistemaVentas.
--
--  ATENCION: este script elimina las tablas si ya existen,
--  por lo que borra todos los datos. Usarlo solo para instalar
--  o reiniciar el sistema.
-- =============================================================

-- Zona horaria de Republica Dominicana para que las fechas de
-- las ventas coincidan con la hora local (el contenedor usa UTC).
DO $$
BEGIN
    EXECUTE format('ALTER DATABASE %I SET timezone TO %L',
                   current_database(), 'America/Santo_Domingo');
END $$;

SET timezone TO 'America/Santo_Domingo';

DROP TABLE IF EXISTS movimientos_inventario CASCADE;
DROP TABLE IF EXISTS detalle_venta CASCADE;
DROP TABLE IF EXISTS ventas CASCADE;
DROP TABLE IF EXISTS productos CASCADE;
DROP TABLE IF EXISTS clientes CASCADE;

-- -------------------------------------------------------------
--  Clientes
-- -------------------------------------------------------------
CREATE TABLE clientes (
    id_cliente INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nombre     VARCHAR(100) NOT NULL,
    telefono   VARCHAR(20),
    correo     VARCHAR(150),

    CONSTRAINT ck_clientes_nombre
        CHECK (btrim(nombre) <> ''),

    CONSTRAINT ck_clientes_correo
        CHECK (correo IS NULL OR correo LIKE '%_@_%._%')
);

-- En Excel los clientes se seleccionan por nombre, por eso el
-- nombre no se puede repetir (sin importar mayusculas).
CREATE UNIQUE INDEX uq_clientes_nombre ON clientes (lower(btrim(nombre)));

-- -------------------------------------------------------------
--  Productos
-- -------------------------------------------------------------
CREATE TABLE productos (
    id_producto INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    nombre      VARCHAR(150) NOT NULL,
    categoria   VARCHAR(100),
    precio      NUMERIC(10,2) NOT NULL,   -- precio final, ITBIS incluido
    stock       INTEGER NOT NULL DEFAULT 0,

    CONSTRAINT ck_productos_nombre CHECK (btrim(nombre) <> ''),
    CONSTRAINT ck_productos_precio CHECK (precio > 0),
    CONSTRAINT ck_productos_stock  CHECK (stock >= 0)
);

CREATE UNIQUE INDEX uq_productos_nombre ON productos (lower(btrim(nombre)));

-- -------------------------------------------------------------
--  Ventas (encabezado de la factura)
-- -------------------------------------------------------------
CREATE TABLE ventas (
    id_venta   INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    id_cliente INTEGER NOT NULL,
    fecha      TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    subtotal   NUMERIC(12,2) NOT NULL DEFAULT 0,
    itbis      NUMERIC(12,2) NOT NULL DEFAULT 0,
    total      NUMERIC(12,2) NOT NULL DEFAULT 0,
    estado     VARCHAR(10) NOT NULL DEFAULT 'EMITIDA',

    CONSTRAINT fk_ventas_cliente
        FOREIGN KEY (id_cliente)
        REFERENCES clientes(id_cliente),

    CONSTRAINT ck_ventas_estado
        CHECK (estado IN ('EMITIDA', 'ANULADA')),

    CONSTRAINT ck_ventas_montos
        CHECK (subtotal >= 0 AND itbis >= 0 AND total = subtotal + itbis)
);

CREATE INDEX ix_ventas_fecha   ON ventas (fecha);
CREATE INDEX ix_ventas_cliente ON ventas (id_cliente);

-- -------------------------------------------------------------
--  Detalle de cada venta
-- -------------------------------------------------------------
CREATE TABLE detalle_venta (
    id_detalle  INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    id_venta    INTEGER NOT NULL,
    id_producto INTEGER NOT NULL,
    cantidad    INTEGER NOT NULL,
    precio      NUMERIC(10,2) NOT NULL,
    subtotal    NUMERIC(12,2) NOT NULL,

    CONSTRAINT fk_detalle_venta
        FOREIGN KEY (id_venta)
        REFERENCES ventas(id_venta)
        ON DELETE CASCADE,

    CONSTRAINT fk_detalle_producto
        FOREIGN KEY (id_producto)
        REFERENCES productos(id_producto),

    CONSTRAINT ck_detalle_cantidad CHECK (cantidad > 0),
    CONSTRAINT ck_detalle_precio   CHECK (precio >= 0),
    CONSTRAINT ck_detalle_subtotal CHECK (subtotal = cantidad * precio)
);

CREATE INDEX ix_detalle_venta    ON detalle_venta (id_venta);
CREATE INDEX ix_detalle_producto ON detalle_venta (id_producto);

-- -------------------------------------------------------------
--  Movimientos de inventario (kardex)
--  Guarda cada entrada y salida de mercancia.
-- -------------------------------------------------------------
CREATE TABLE movimientos_inventario (
    id_movimiento    INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    id_producto      INTEGER NOT NULL,
    fecha            TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    tipo             VARCHAR(10) NOT NULL,
    cantidad         INTEGER NOT NULL,
    stock_resultante INTEGER NOT NULL,
    referencia       VARCHAR(150),

    CONSTRAINT fk_movimiento_producto
        FOREIGN KEY (id_producto)
        REFERENCES productos(id_producto),

    CONSTRAINT ck_movimiento_tipo
        CHECK (tipo IN ('INICIAL', 'ENTRADA', 'SALIDA', 'ANULACION')),

    -- Las entradas suman (cantidad positiva) y las salidas restan
    -- (cantidad negativa).
    CONSTRAINT ck_movimiento_cantidad CHECK (cantidad <> 0)
);

CREATE INDEX ix_movimientos_producto ON movimientos_inventario (id_producto, fecha);
