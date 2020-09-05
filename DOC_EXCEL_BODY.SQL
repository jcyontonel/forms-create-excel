PACKAGE BODY DOC_EXCEL IS

	--Constructor for type TITULO
	function NEW_TITULO(P_TITULO VARCHAR2, P_TIPO VARCHAR2) return R_TITULO is
	  OBJ_TITULO R_TITULO;
	begin
	  OBJ_TITULO.TITULO := P_TITULO;
	  OBJ_TITULO.TIPO := P_TIPO;
	  return(obj_titulo);
	end;
	
	PROCEDURE INICIAR IS
	BEGIN
		application := ole2.create_obj('EXCEL.APPLICATION');
		ole2.set_property(application, 'VISIBLE', 'FALSE');
		workbooks   := ole2.get_obj_property(application, 'WORKBOOKS');
		--Adiciona un LIBRO (DOCUMENTO EXCEL)
		workbook    := ole2.get_obj_property(workbooks, 'ADD');
		--lee el handle de la coleccion de hojas
		worksheets  := ole2.get_obj_property(workbook, 'WORKSHEETS');
	END;
	
	PROCEDURE CERRAR IS
	BEGIN
	  ole2.set_property(application, 'Visible', 'True');
	  -- RELEASE THE OLE OBJECTS
	  ole2.release_obj(worksheet);
	  ole2.release_obj(worksheets);
	  ole2.release_obj(workbook);
	  ole2.release_obj(workbooks);
	  ole2.release_obj(application);
	END;
	
	PROCEDURE AUTOAJUSTAR IS
	BEGIN
	  RANGE     := ole2.get_obj_property(worksheet, 'USEDRANGE');
	  range_col := ole2.get_obj_property(RANGE, 'COLUMNS');
	  ole2.invoke(range_col, 'AUTOFIT');
	  ole2.release_obj(RANGE);
	  ole2.release_obj(range_col);
	END;

	
	PROCEDURE ABRIR_HOJA(numero_hoja NUMBER, x_nombre VARCHAR2) IS
	BEGIN
		-- SI EL NUMERO DE HOJA EXISTE LA ABRIMOS
		IF (numero_hoja <= ole2.get_num_property(worksheets, 'count')) THEN
			args := ole2.create_arglist;
			ole2.add_arg(args, numero_hoja);
			worksheet := ole2.get_obj_property(worksheets, 'ITEM', args);
			ole2.set_property(worksheet, 'Name', x_nombre);
			ole2.destroy_arglist(args);
			
		-- SINO AGREGAMOS UNA NUEVA HOJA
		ELSE
			worksheet := OLE2.GET_OBJ_PROPERTY(worksheets, 'Add');
			ole2.set_property(worksheet, 'Name', x_nombre);
		END IF;
	END;
	
	PROCEDURE SET_CELDA_TITULO(fila NUMBER, columna NUMBER, valor VARCHAR2) IS
	BEGIN
		PRINT_CELDA(fila, columna, valor, 'GENERAL', 'T', COLORS(NULL));
	END;
	
	PROCEDURE SET_CELDA_RESALTADA(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL') IS
	BEGIN
		PRINT_CELDA(fila, columna, valor, tipo_dato, 'R', COLORS(NULL));
	END;
	
	PROCEDURE SET_CELDA_COLOR(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL', COLOR VARCHAR2) IS
	BEGIN
		PRINT_CELDA(fila, columna, valor, tipo_dato, NULL, COLORS(COLOR));
	END;
	
	PROCEDURE SET_CELDA(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL') IS
	BEGIN
		PRINT_CELDA(fila, columna, valor, tipo_dato, NULL, COLORS(NULL));
	END;
	
	PROCEDURE PRINT_CELDA(fila NUMBER, columna NUMBER, valor VARCHAR2, 
											tipo_dato VARCHAR2 DEFAULT 'GENERAL', tipo_celda VARCHAR2, COLOR_CELDA NUMBER) IS
  BEGIN
		/* DEFINIMOS LA CELDA */
		args := ole2.create_arglist;
		ole2.add_arg(args, fila);
		ole2.add_arg(args, columna);
		cell := ole2.get_obj_property(worksheet, 'CELLS', args);
		ole2.destroy_arglist(args);
		-- PROPIEDADES 
		ole2.set_property(cell, 'NUMBERFORMAT', tipo_dato);
		ole2.set_property(cell, 'VALUE', valor);
		/* DEFINIMOS EL FORMATO SEGÚN TIPO DE CELDA ELEGIDA */
		font  := OLE2.GET_OBJ_PROPERTY(cell, 'Font');
		OLE2.SET_PROPERTY(font, 'Name', 'Calibri');
		ole2.set_property(font, 'Size', 10);
		--- COLOR DE CELDA
		IF COLOR_CELDA IS NOT NULL THEN
			colour:=ole2.get_obj_property(cell, 'Interior');
			ole2.set_property(colour, 'ColorIndex', COLOR_CELDA);
		END IF;
		-- TIPO DE FUENTE
		IF tipo_celda = 'T' THEN
			ole2.set_property(font, 'BOLD', TRUE);
		  ole2.set_property(font, 'Size', 10);
		ELSIF tipo_celda = 'R' THEN
			ole2.set_property(font, 'Size', 12);
			ole2.set_property(font, 'BOLD', TRUE);
		  OLE2.SET_PROPERTY(font, 'ColorIndex', 5);-- 3Red  5 azul
		END IF;
		/* DESTRUIMOS LOS OBJETOS CREADOS */
		ole2.release_obj(font);
		ole2.release_obj(cell);
  END;
  
	PROCEDURE IMPRIMIR_TITULOS (nfila NUMERIC, X_TITULOS ARRAY_TITULOS)IS
	BEGIN
		------- IMPRIMIMOS LOS ENCABEZADOS DE COLUMNAS -------------------------------------
		FOR ncolumn IN 1 .. X_TITULOS.COUNT LOOP
		   SET_CELDA_TITULO(nfila, ncolumn, X_TITULOS(ncolumn));
		END LOOP;
	END;
	
	
	--functions
	FUNCTION COLORS(X_NAME VARCHAR2) RETURN NUMBER IS
		X_COLOR NUMBER;
	BEGIN
		IF X_NAME = 'ROJO' THEN
			 RETURN ROJO;
		ELSIF X_NAME = 'VERDE' THEN
			 RETURN VERDE;
		ELSIF X_NAME = 'AZUL' THEN
			 RETURN AZUL;
		ELSIF X_NAME = 'AMARILLO' THEN
			 RETURN AMARILLO;
		ELSIF X_NAME = 'NARANJA' THEN
			 RETURN NARANJA;
		ELSIF X_NAME = 'CELESTE' THEN
			 RETURN CELESTE;
		ELSIF X_NAME = 'ROSADO' THEN
			 RETURN ROSADO;
		ELSE 
			RETURN NULL;
		END IF;
	END;
END;
