-- =============================================================
--  Sistema de Ventas e Inventario
--  02_funciones_y_vistas.sql
--  Logica de negocio (funciones), triggers y vistas para reportes.
--
--  Las reglas importantes (validar stock, descontar inventario,
--  calcular ITBIS, anular ventas) viven en la base de datos para
--  que se ejecuten dentro de una sola transaccion: o se guarda
--  todo, o no se guarda nada.
-- =============================================================

-- -------------------------------------------------------------
--  Tasa de ITBIS usada por el sistema (18 %)
--  Los precios de los productos YA INCLUYEN el ITBIS.
--  Al vender, el total se desglosa:
--      subtotal = total / 1.18      itbis = total - subtotal
--  Ejemplo: producto de 1,180.00 -> subtotal 1,000.00 + ITBIS 180.00
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION tasa_itbis()
RETURNS NUMERIC
LANGUAGE sql IMMUTABLE
AS $$ SELECT 0.18::NUMERIC $$;

-- -------------------------------------------------------------
--  Numero de factura con formato FAC-000001
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION numero_factura(p_id_venta INTEGER)
RETURNS TEXT
LANGUAGE sql IMMUTABLE
AS $$ SELECT 'FAC-' || lpad(p_id_venta::TEXT, 6, '0') $$;

-- -------------------------------------------------------------
--  registrar_venta
--  Registra una venta completa con todas sus lineas.
--
--  Parametros:
--    p_id_cliente  : cliente que compra
--    p_productos   : arreglo con los id de producto  {1,2,3}
--    p_cantidades  : arreglo con las cantidades       {2,1,5}
--
--  Devuelve el id_venta generado. Los precios incluyen ITBIS.
--  Si algun producto no tiene stock suficiente se cancela
--  toda la venta (no se guarda nada).
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION registrar_venta(
    p_id_cliente INTEGER,
    p_productos  INTEGER[],
    p_cantidades INTEGER[]
)
RETURNS INTEGER
LANGUAGE plpgsql
AS $$
DECLARE
    v_id_venta   INTEGER;
    v_producto   productos%ROWTYPE;
    v_cantidad   INTEGER;
    v_total      NUMERIC(12,2) := 0;
    v_subtotal   NUMERIC(12,2);
    v_itbis      NUMERIC(12,2);
    v_nuevo      INTEGER;
    i            INTEGER;
BEGIN
    -- Validaciones generales
    IF NOT EXISTS (SELECT 1 FROM clientes WHERE id_cliente = p_id_cliente) THEN
        RAISE EXCEPTION 'El cliente con ID % no existe.', p_id_cliente;
    END IF;

    IF p_productos IS NULL OR cardinality(p_productos) = 0 THEN
        RAISE EXCEPTION 'La venta debe tener al menos un producto.';
    END IF;

    IF p_cantidades IS NULL
       OR cardinality(p_productos) <> cardinality(p_cantidades) THEN
        RAISE EXCEPTION 'La lista de productos y la de cantidades no coinciden.';
    END IF;

    -- Encabezado de la venta (los montos se calculan despues)
    INSERT INTO ventas (id_cliente)
    VALUES (p_id_cliente)
    RETURNING id_venta INTO v_id_venta;

    -- Lineas de la venta
    FOR i IN 1 .. cardinality(p_productos) LOOP
        v_cantidad := p_cantidades[i];

        IF v_cantidad IS NULL OR v_cantidad <= 0 THEN
            RAISE EXCEPTION 'La cantidad de la linea % debe ser mayor que cero.', i;
        END IF;

        -- FOR UPDATE bloquea el producto para que dos ventas al
        -- mismo tiempo no vendan el mismo stock.
        SELECT * INTO v_producto
        FROM productos
        WHERE id_producto = p_productos[i]
        FOR UPDATE;

        IF NOT FOUND THEN
            RAISE EXCEPTION 'El producto con ID % no existe.', p_productos[i];
        END IF;

        IF v_producto.stock < v_cantidad THEN
            RAISE EXCEPTION 'Stock insuficiente para "%". Disponible: %, solicitado: %.',
                v_producto.nombre, v_producto.stock, v_cantidad;
        END IF;

        INSERT INTO detalle_venta (id_venta, id_producto, cantidad, precio, subtotal)
        VALUES (v_id_venta, v_producto.id_producto, v_cantidad,
                v_producto.precio, v_cantidad * v_producto.precio);

        UPDATE productos
        SET stock = stock - v_cantidad
        WHERE id_producto = v_producto.id_producto
        RETURNING stock INTO v_nuevo;

        INSERT INTO movimientos_inventario
            (id_producto, tipo, cantidad, stock_resultante, referencia)
        VALUES
            (v_producto.id_producto, 'SALIDA', -v_cantidad, v_nuevo,
             'Venta ' || numero_factura(v_id_venta));

        v_total := v_total + v_cantidad * v_producto.precio;
    END LOOP;

    -- Los precios incluyen ITBIS: se desglosa a partir del total
    v_subtotal := round(v_total / (1 + tasa_itbis()), 2);
    v_itbis    := v_total - v_subtotal;

    UPDATE ventas
    SET subtotal = v_subtotal,
        itbis    = v_itbis,
        total    = v_total
    WHERE id_venta = v_id_venta;

    RETURN v_id_venta;
END;
$$;

-- -------------------------------------------------------------
--  anular_venta
--  Marca la venta como ANULADA y devuelve los productos al
--  inventario. Una venta anulada no se puede volver a anular.
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION anular_venta(p_id_venta INTEGER)
RETURNS VOID
LANGUAGE plpgsql
AS $$
DECLARE
    v_estado  VARCHAR(10);
    v_linea   RECORD;
    v_nuevo   INTEGER;
BEGIN
    SELECT estado INTO v_estado
    FROM ventas
    WHERE id_venta = p_id_venta
    FOR UPDATE;

    IF NOT FOUND THEN
        RAISE EXCEPTION 'La venta % no existe.', numero_factura(p_id_venta);
    END IF;

    IF v_estado = 'ANULADA' THEN
        RAISE EXCEPTION 'La venta % ya estaba anulada.', numero_factura(p_id_venta);
    END IF;

    FOR v_linea IN
        SELECT id_producto, cantidad
        FROM detalle_venta
        WHERE id_venta = p_id_venta
        ORDER BY id_producto
    LOOP
        UPDATE productos
        SET stock = stock + v_linea.cantidad
        WHERE id_producto = v_linea.id_producto
        RETURNING stock INTO v_nuevo;

        INSERT INTO movimientos_inventario
            (id_producto, tipo, cantidad, stock_resultante, referencia)
        VALUES
            (v_linea.id_producto, 'ANULACION', v_linea.cantidad, v_nuevo,
             'Anulacion ' || numero_factura(p_id_venta));
    END LOOP;

    UPDATE ventas
    SET estado = 'ANULADA'
    WHERE id_venta = p_id_venta;
END;
$$;

-- -------------------------------------------------------------
--  entrada_inventario
--  Suma mercancia al stock de un producto (compra, reposicion).
--  Devuelve el nuevo stock.
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION entrada_inventario(
    p_id_producto INTEGER,
    p_cantidad    INTEGER,
    p_referencia  VARCHAR DEFAULT NULL
)
RETURNS INTEGER
LANGUAGE plpgsql
AS $$
DECLARE
    v_nuevo INTEGER;
BEGIN
    IF p_cantidad IS NULL OR p_cantidad <= 0 THEN
        RAISE EXCEPTION 'La cantidad de entrada debe ser mayor que cero.';
    END IF;

    UPDATE productos
    SET stock = stock + p_cantidad
    WHERE id_producto = p_id_producto
    RETURNING stock INTO v_nuevo;

    IF NOT FOUND THEN
        RAISE EXCEPTION 'El producto con ID % no existe.', p_id_producto;
    END IF;

    INSERT INTO movimientos_inventario
        (id_producto, tipo, cantidad, stock_resultante, referencia)
    VALUES
        (p_id_producto, 'ENTRADA', p_cantidad, v_nuevo,
         COALESCE(NULLIF(btrim(p_referencia), ''), 'Entrada de inventario'));

    RETURN v_nuevo;
END;
$$;

-- -------------------------------------------------------------
--  Trigger: al crear un producto con stock se registra el
--  movimiento INICIAL en el kardex.
-- -------------------------------------------------------------
CREATE OR REPLACE FUNCTION fn_producto_stock_inicial()
RETURNS TRIGGER
LANGUAGE plpgsql
AS $$
BEGIN
    IF NEW.stock > 0 THEN
        INSERT INTO movimientos_inventario
            (id_producto, tipo, cantidad, stock_resultante, referencia)
        VALUES
            (NEW.id_producto, 'INICIAL', NEW.stock, NEW.stock, 'Stock inicial');
    END IF;
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS tr_producto_stock_inicial ON productos;

CREATE TRIGGER tr_producto_stock_inicial
AFTER INSERT ON productos
FOR EACH ROW
EXECUTE FUNCTION fn_producto_stock_inicial();

-- =============================================================
--  VISTAS
-- =============================================================

-- Listado de ventas para el historial
CREATE OR REPLACE VIEW v_ventas_resumen AS
SELECT
    v.id_venta,
    numero_factura(v.id_venta)                    AS numero,
    v.fecha,
    c.nombre                                      AS cliente,
    COALESCE((SELECT SUM(d.cantidad)
              FROM detalle_venta d
              WHERE d.id_venta = v.id_venta), 0)::INTEGER AS articulos,
    v.subtotal,
    v.itbis,
    v.total,
    v.estado
FROM ventas v
INNER JOIN clientes c ON c.id_cliente = v.id_cliente;

-- Encabezado de la factura
CREATE OR REPLACE VIEW v_factura_encabezado AS
SELECT
    v.id_venta,
    numero_factura(v.id_venta) AS numero,
    v.fecha,
    v.estado,
    c.id_cliente,
    c.nombre   AS cliente,
    c.telefono,
    c.correo,
    v.subtotal,
    v.itbis,
    v.total
FROM ventas v
INNER JOIN clientes c ON c.id_cliente = v.id_cliente;

-- Lineas de la factura
CREATE OR REPLACE VIEW v_factura_detalle AS
SELECT
    d.id_venta,
    d.id_detalle,
    p.id_producto,
    p.nombre AS producto,
    d.cantidad,
    d.precio,
    d.subtotal
FROM detalle_venta d
INNER JOIN productos p ON p.id_producto = d.id_producto;

-- Productos con 5 unidades o menos
CREATE OR REPLACE VIEW v_stock_bajo AS
SELECT id_producto, nombre, categoria, stock
FROM productos
WHERE stock <= 5
ORDER BY stock, nombre;

-- Kardex con el nombre del producto
CREATE OR REPLACE VIEW v_kardex AS
SELECT
    m.id_movimiento,
    m.fecha,
    p.nombre AS producto,
    m.tipo,
    m.cantidad,
    m.stock_resultante,
    m.referencia
FROM movimientos_inventario m
INNER JOIN productos p ON p.id_producto = m.id_producto;

-- Indicadores para el menu principal
CREATE OR REPLACE VIEW v_indicadores AS
SELECT
    (SELECT COUNT(*) FROM clientes)::INTEGER  AS total_clientes,
    (SELECT COUNT(*) FROM productos)::INTEGER AS total_productos,
    (SELECT COUNT(*) FROM v_stock_bajo)::INTEGER AS productos_stock_bajo,
    (SELECT COUNT(*) FROM ventas
      WHERE estado = 'EMITIDA'
        AND fecha::DATE = CURRENT_DATE)::INTEGER AS ventas_hoy,
    (SELECT COALESCE(SUM(total), 0) FROM ventas
      WHERE estado = 'EMITIDA'
        AND fecha::DATE = CURRENT_DATE) AS total_hoy,
    (SELECT COALESCE(SUM(total), 0) FROM ventas
      WHERE estado = 'EMITIDA'
        AND date_trunc('month', fecha) = date_trunc('month', CURRENT_DATE)) AS total_mes;

-- =============================================================
--  FUNCIONES DE REPORTES (por rango de fechas)
--  Solo toman en cuenta ventas EMITIDAS (no anuladas).
-- =============================================================

CREATE OR REPLACE FUNCTION reporte_resumen(p_desde DATE, p_hasta DATE)
RETURNS TABLE (
    cantidad_ventas  INTEGER,
    subtotal         NUMERIC,
    itbis            NUMERIC,
    total            NUMERIC,
    ticket_promedio  NUMERIC,
    ventas_anuladas  INTEGER
)
LANGUAGE sql STABLE
AS $$
    SELECT
        COUNT(*) FILTER (WHERE v.estado = 'EMITIDA')::INTEGER,
        COALESCE(SUM(v.subtotal) FILTER (WHERE v.estado = 'EMITIDA'), 0),
        COALESCE(SUM(v.itbis)    FILTER (WHERE v.estado = 'EMITIDA'), 0),
        COALESCE(SUM(v.total)    FILTER (WHERE v.estado = 'EMITIDA'), 0),
        COALESCE(round(AVG(v.total) FILTER (WHERE v.estado = 'EMITIDA'), 2), 0),
        COUNT(*) FILTER (WHERE v.estado = 'ANULADA')::INTEGER
    FROM ventas v
    WHERE v.fecha::DATE BETWEEN p_desde AND p_hasta;
$$;

CREATE OR REPLACE FUNCTION reporte_ventas_por_dia(p_desde DATE, p_hasta DATE)
RETURNS TABLE (
    dia              DATE,
    cantidad_ventas  INTEGER,
    total            NUMERIC
)
LANGUAGE sql STABLE
AS $$
    SELECT
        v.fecha::DATE,
        COUNT(*)::INTEGER,
        SUM(v.total)
    FROM ventas v
    WHERE v.estado = 'EMITIDA'
      AND v.fecha::DATE BETWEEN p_desde AND p_hasta
    GROUP BY v.fecha::DATE
    ORDER BY v.fecha::DATE;
$$;

CREATE OR REPLACE FUNCTION reporte_productos_mas_vendidos(
    p_desde  DATE,
    p_hasta  DATE,
    p_limite INTEGER DEFAULT 10
)
RETURNS TABLE (
    producto         VARCHAR,
    cantidad_vendida INTEGER,
    ingresos         NUMERIC
)
LANGUAGE sql STABLE
AS $$
    SELECT
        p.nombre,
        SUM(d.cantidad)::INTEGER,
        SUM(d.subtotal)
    FROM detalle_venta d
    INNER JOIN ventas v    ON v.id_venta = d.id_venta
    INNER JOIN productos p ON p.id_producto = d.id_producto
    WHERE v.estado = 'EMITIDA'
      AND v.fecha::DATE BETWEEN p_desde AND p_hasta
    GROUP BY p.nombre
    ORDER BY SUM(d.cantidad) DESC, SUM(d.subtotal) DESC
    LIMIT p_limite;
$$;
