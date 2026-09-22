INSERT INTO clientes (nombre, telefono, correo)
VALUES
('Edgar Fortuna', '809-555-1001', 'edgar@example.com'),
('Carlos Pérez', '809-555-1002', 'carlos@example.com');

SELECT * FROM clientes;

INSERT INTO productos (nombre, categoria, precio, stock)
VALUES
('Teclado Logitech', 'Periféricos', 1500.00, 10),
('Mouse inalámbrico', 'Periféricos', 850.00, 15),
('Monitor 24 pulgadas', 'Monitores', 8500.00, 5);

SELECT * FROM productos;

INSERT INTO ventas (id_cliente, total)
VALUES (1, 3200.00);

SELECT * FROM ventas;

UPDATE ventas
SET  total = 3850.00
WHERE id_venta = 1;

INSERT INTO detalle_venta
    (id_venta, id_producto, cantidad, precio, subtotal)
VALUES
    (1, 1, 2, 1500.00, 3000.00),
    (1, 2, 1, 850.00, 850.00);
    
    SELECT 
    	v.id_venta,
    	c.nombre AS cliente,
    	p.nombre AS producto,
    	dv.cantidad,
    	dv.precio,
    	dv.subtotal,
    	v.fecha
    	FROM ventas AS v
    	INNER JOIN clientes AS c
    		ON v.id_cliente = c.id_cliente
    	INNER JOIN detalle_venta AS dv 
    		ON v.id_venta = dv.id_venta
    	INNER JOIN productos AS p 
    		ON dv.id_producto = p.id_producto
    	WHERE v.id_venta =1;
    	
    	