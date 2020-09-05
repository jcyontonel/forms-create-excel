PACKAGE DOC_EXCEL IS
	-- DECLARE THE OLE OBJECTS
	application                  ole2.obj_type;
	workbooks                    ole2.obj_type;
	workbook                     ole2.obj_type;
	worksheets                   ole2.obj_type;
	worksheet                    ole2.obj_type;
	cell                         ole2.obj_type;
	font  					     	  		ole2.obj_type;
	RANGE                        ole2.obj_type;
	range_col                    ole2.obj_type;
	args                         ole2.list_type;
	colour 												ole2.obj_type;
	
	-- CONSTANTES TIPOS DE DATO DE LAS CELDAS
	TEXTO 		VARCHAR2(15) 	:= '@';             
	DECIMALL 	VARCHAR2(15)	:= '###,###,##0.00';
	NUMERO 	VARCHAR2(15)	:= 'NUMERICO';
	ENTERO 	VARCHAR2(15)	:= '###,###,##0';
  FECHA 		VARCHAR2(15)	:= 'DD/MM/YYYY';      
  GENERAL 	VARCHAR2(15)	:= 'GENERAL';
  
  -- COLORES
  ROJO 		NUMBER := 3;
  VERDE 	NUMBER := 43;
  AZUL		NUMBER := 23;
  AMARILLO NUMBER := 44;
  NARANJA NUMBER := 45;
  CELESTE NUMBER:= 33;
  ROSADO NUMBER := 38;
  BLANCO NUMBER := 2;

  
	-- DECLARAMOS TIPO DE DATO
	type R_TITULO is record(TITULO VARCHAR2(100), TIPO VARCHAR2(20));
	--type ARRAY_TITULOS is table of R_TITULO;
	
	type ARRAY_TITULOS IS TABLE OF VARCHAR2(50) INDEX BY BINARY_INTEGER; 
	TITULOS ARRAY_TITULOS;
	
	function NEW_TITULO(P_TITULO VARCHAR2, P_TIPO VARCHAR2) return R_TITULO;
	
	PROCEDURE INICIAR;
	PROCEDURE CERRAR;
	PROCEDURE AUTOAJUSTAR;
	PROCEDURE ABRIR_HOJA(numero_hoja NUMBER, x_nombre VARCHAR2 DEFAULT NULL);
	PROCEDURE SET_CELDA_COLOR(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL', COLOR VARCHAR2);
	PROCEDURE SET_CELDA_TITULO(fila NUMBER, columna NUMBER, valor VARCHAR2);
	PROCEDURE SET_CELDA_RESALTADA(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL');
	PROCEDURE SET_CELDA(fila NUMBER, columna NUMBER, valor VARCHAR2, tipo_dato VARCHAR2 DEFAULT 'GENERAL');
	PROCEDURE PRINT_CELDA(fila NUMBER, columna NUMBER, valor VARCHAR2, 
											tipo_dato VARCHAR2 DEFAULT 'GENERAL', tipo_celda VARCHAR2, COLOR_CELDA NUMBER);
	PROCEDURE IMPRIMIR_TITULOS (nfila NUMERIC, X_TITULOS ARRAY_TITULOS);
	FUNCTION COLORS(X_NAME VARCHAR2) RETURN NUMBER;
END;
