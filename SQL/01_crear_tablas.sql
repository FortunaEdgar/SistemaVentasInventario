CREATE TABLE clientes (
	id_cliente INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
	nombre VARCHAR(100) not NULL,
	telefono VARCHAR(20),
	correo VARCHAR(150)
);

CREATE TABLE productos (
	id_producto INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
	nombre VARCHAR(150) NOT NULL,
	categoria VARCHAR(100),
	precio NUMERIC(10,2) NOT NULL,
	stock INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE ventas (
	id_venta INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
	id_cliente INTEGER NOT NULL,
	fecha TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
	total NUMERIC(10,2) NOT NULL DEFAULT 0,
	
	CONSTRAINT fk_ventas_cliente
		FOREIGN KEY (id_cliente)
		REFERENCES clientes(id_cliente)
);

CREATE TABLE detalle_venta (
	id_detalle INTEGER GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
	id_venta INTEGER NOT NULL,
	id_producto INTEGER NOT NULL,
	cantidad INTEGER NOT NULL,
	precio NUMERIC(10,2) NOT NULL,
	subtotal NUMERIC(10,2) NOT NULL,
	
	CONSTRAINT fk_detalle_venta
		FOREIGN KEY (id_venta)
		REFERENCES ventas(id_venta),
		
		CONSTRAINT fk_detalle_producto
			FOREIGN KEY (id_producto)
			REFERENCES 	productos(id_producto)
);