-- =============================================================
--  Sistema de Ventas e Inventario
--  03_datos_prueba.sql
--  Datos de ejemplo para probar el sistema.
--  Las ventas se registran con la funcion registrar_venta(),
--  asi el stock y el kardex quedan correctos.
-- =============================================================

SET timezone TO 'America/Santo_Domingo';

INSERT INTO clientes (nombre, telefono, correo)
VALUES
    ('Consumidor Final', NULL,           NULL),
    ('Edgar Fortuna',    '809-555-1001', 'edgar@example.com'),
    ('Carlos Pérez',     '809-555-1002', 'carlos@example.com'),
    ('Maria Rodriguez',  '809-555-9025', 'maria@example.com'),
    ('Hector Yan',       '829-555-3410', NULL);

INSERT INTO productos (nombre, categoria, precio, stock)
VALUES
    ('Teclado Logitech',      'Periféricos',     1500.00, 20),
    ('Mouse inalámbrico',     'Periféricos',      850.00, 25),
    ('Monitor 24 pulgadas',   'Monitores',       8500.00,  8),
    ('Audifonos Bluetooth',   'Audio',           1200.00, 12),
    ('Memoria USB 64GB',      'Almacenamiento',   650.00, 30),
    ('Disco SSD 500GB',       'Almacenamiento',  3200.00, 10),
    ('Cable HDMI 2m',         'Accesorios',       350.00,  4),
    ('Webcam HD',             'Periféricos',     2100.00,  6);

-- Funcion auxiliar solo para los datos de prueba
CREATE OR REPLACE FUNCTION fecha_prueba(p_id_venta INTEGER, p_dias INTEGER)
RETURNS VOID LANGUAGE sql AS $$
    UPDATE ventas
    SET fecha = CURRENT_TIMESTAMP - make_interval(days => p_dias)
    WHERE id_venta = p_id_venta;

    UPDATE movimientos_inventario
    SET fecha = CURRENT_TIMESTAMP - make_interval(days => p_dias)
    WHERE referencia = 'Venta ' || numero_factura(p_id_venta);
$$;

-- Ventas de ejemplo (cliente, productos, cantidades)
DO $$
DECLARE
    v_id INTEGER;
BEGIN
    -- Despues de cada venta se cambia la fecha para simular dias
    -- anteriores (la venta y sus movimientos de inventario).
    v_id := registrar_venta(2, ARRAY[1, 2], ARRAY[2, 1]);
    PERFORM fecha_prueba(v_id, 6);

    v_id := registrar_venta(3, ARRAY[3], ARRAY[1]);
    PERFORM fecha_prueba(v_id, 4);

    v_id := registrar_venta(4, ARRAY[4, 5, 7], ARRAY[2, 3, 1]);
    PERFORM fecha_prueba(v_id, 2);

    v_id := registrar_venta(1, ARRAY[5, 2], ARRAY[2, 1]);
    PERFORM fecha_prueba(v_id, 1);

    v_id := registrar_venta(5, ARRAY[6, 1], ARRAY[1, 1]);

    -- Ejemplo de venta anulada (el stock se devuelve)
    v_id := registrar_venta(1, ARRAY[8], ARRAY[1]);
    PERFORM anular_venta(v_id);
END $$;

DROP FUNCTION fecha_prueba(INTEGER, INTEGER);

-- Consultas de verificacion ----------------------------------
-- SELECT * FROM v_ventas_resumen ORDER BY id_venta;
-- SELECT * FROM v_factura_detalle WHERE id_venta = 1;
-- SELECT * FROM v_kardex ORDER BY id_movimiento;
-- SELECT * FROM v_indicadores;
