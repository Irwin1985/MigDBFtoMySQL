&&*-----------------------------------------------------------------------------------------------*
&&*... Intercepta y maneja los errores del Sistema en Tiempo de Ejecución
&&*-----------------------------------------------------------------------------------------------*
PROCEDURE MANEJAERROR
	PARAMETERS PROGRAMA,LINEA,NUMERROR,MENSAJE

	DO CASE
	CASE NUMERROR = 125  && Impresora no esta encendida
		MESSAGEBOX('Impresora <<NO>> preparada')
		RETURN
	CASE NUMERROR = 5 && Registro fuera de rango
	CASE NUMERROR = 39 &&Error Desbordamiento Númerico
		NROERROR = 39
	CASE NUMERROR = 1307 &&Error División por cero
		NROERROR = 1307
	CASE NUMERROR = 2091 &&Tabla Dañada o Corrupta
	OTHERWISE
		DO ERRORTXT WITH PROGRAMA,LINEA,NUMERROR,MENSAJE
		IF MESSAGEBOX("ERROR "+STR(NUMERROR,2)+" "+MENSAJE+"¿ Desea Salir ?",3+32,"¿Continuar con la Aplicación?") = 6
			CANCEL
			CLOSE ALL
			QUIT
		ELSE &&MESSAGEBOX("ERROR "+STR(NUMERROR,2)+" "+MENSAJE+"¿ Desea Salir ?",3+32,"¿Continuar con la Aplicación?") = 7
		ENDIF &&MESSAGEBOX("ERROR "+STR(NUMERROR,2)+" "+MENSAJE+"¿ Desea Salir ?",3+32,"¿Continuar con la Aplicación?") = 7
	ENDCASE
	RETURN
&&*-----------------------------------------------------------------------------------------------*
&&*...
&&*-----------------------------------------------------------------------------------------------*
PROCEDURE ERRORTXT
	PARAMETERS PROCNOM,LINEA,ERRCOD,ERRMESS

	IF (ERRCOD = 2002) OR (ERRCOD = 2004)
		ARC = 'MEMORIA.REF'
	ELSE &&(ERRCOD = 2002) OR (ERRCOD = 2004)
		ARC = 'ERR'+SUBSTR(SYS(3),1,5)+'.TXT'
	ENDIF &&(ERRCOD = 2002) OR (ERRCOD = 2004)

	ARC = gcDirTemp+ARC
	OLDPRINT = SET('PRINT')
	SET SAFETY OFF
	SET PRINT OFF
	SET ALTERNATE TO (ARC)
	SET ALTERNATE ON
	SET CONSOLE OFF
	? 'ERROR :  VACIADO DE LA INFORMACION DEL SISTEMA'
	?
	? '                 FECHA:  ' +DTOC(DATE())
	? '                  HORA:  ' +TIME()
	? '   NOMBRE DEL PROGRAMA:  ' +PROCNOM
	? '    LINEA DEL PROGRAMA:  ' +LTRIM(STR(LINEA))
	? '      MENSAJE DE ERROR:  ' +ERRMESS
	? '       CODIGO DE ERROR:  ' +STR(ERRCOD)
	?
	? '    MEMORIA DISPONIBLE:  ' +SYS(12)
	? '         MEMORIA USADA:  ' +SYS(23)
	? ' LIMITE DE MEMORIA EMS:  ' +SYS(24)
	? '    MEMORIA DE USUARIO:  ' +SYS(1016)
	? '      CONFIG.SYS FILES:  ' +SYS(2010)
	?
	? '================================='+;
		'=================='
	?
	? '                   LISTA DE STATUS'
	?
	LIST STATUS
	?
	? '===================================================='+;
		'=================='
	?
	? '                  LISTA DE MEMORIA'
	?
	LIST MEMORY
	?

	? '===================================================='+;
		'=================='
	?
	? '                   JERARQUÍA DE LLAMADAS'
	?
	I=1
	DO WHILE LEN(SYS(16,I)) > 0
		? LTRIM(STR(I)) +' -- '
		?? SYS(16,I)
		I=I +1
		?
	ENDDO
	? '===================================================='+;
		'=================='
	?
	? '                  LISTA DE OBJETOS'
	?
	LIST OBJECTS
	?
	? '===================================================='+;
		'=================='
	?
	? '               STATUS DE LAS AREAS DE TRABAJO'
	?
	? 'AREA ACTUAL '+ALIAS()
	?
	? 'AREA TRAB    NOMBRE DE DBF     NUMERO REGIST      '+;
		'INDICE MAEST'
	? '---------    ------------      -------------      '+;
		'----------'
	FOR I=1 TO 25
		IF NOT EMPTY(ALIAS(I))
			SELECT(I)
			? '  '  +LTRIM(STR(I))
			?? ALIAS() AT 15
			?? LTRIM(STR(RECNO())) AT 35
			?? SYS(21) AT 55
		ELSE &&NOT EMPTY(ALIAS(I))
		ENDIF &&NOT EMPTY(ALIAS(I))
	ENDFOR &&I=1 TO 25
	?
	? '===================================================='+;
		'=================='
	?
	? '              DETALLES DEL AREA ABIERTA'
	?
	? '----------------------------------------------------'
	FOR I=1 TO 25
		IF NOT EMPTY(ALIAS(I))
			?
			SELECT(I)
			LIST STRUCTURE
			?
			? 'CONTENIDO DE REGISTRO ACTUAL DE ' +ALIAS(I)
			?
			? 'NOM. CAMPO   CONTENIDO'
			? '----------   ---------'
			FOR J=1 TO FCOUNT()
				_F_NAME=FIELD(J)
				? _F_NAME
				?? &_F_NAME AT 12
			ENDFOR
			?
			? '----------------------------------------------------'
		ELSE &&NOT EMPTY(ALIAS(I))
		ENDIF &&NOT EMPTY(ALIAS(I))
	ENDFOR &&I=1 TO 25

	SET CONSOLE ON
	SET ALTERNATE OFF
	SET ALTERNATE TO
	SET PRINT &OLDPRINT
	SET SAFETY ON

	RETURN
&&*-----------------------------------------------------------------------------------------------*
&&*... Carga los Datos de conexión para el servidor MySQL
&&*-----------------------------------------------------------------------------------------------*
PROCEDURE CargaDatSql()
	PRIVATE pcArcMySQL
	pcArcMySQL = ALLTRIM(gcPath) + "MySQL.INI"
	IF !FILE(pcArcMySQL)
		luValor = ADDBS(SYS(5)+SYS(2003))
		gcPath = ALLTRIM(luValor)
		WriteIniEntry(luValor,"GENERAL","CURDIR", pcArcMySQL)

		luValor = ADDBS(gcPath+'DATA')
		WriteIniEntry(luValor,"GENERAL","DIRDAT", pcArcMySQL)

		luValor = ADDBS(gcPath+'TEMP')
		WriteIniEntry(luValor,"GENERAL","DIRTEMP", pcArcMySQL)

		luValor = ADDBS(gcPath+'REP')
		WriteIniEntry(luValor,"GENERAL","DIRREP", pcArcMySQL)

		luValor = "3306"
		WriteIniEntry(luValor,"MYSQL","PORT", pcArcMySQL)

		luValor = "Nombre DB"
		WriteIniEntry(luValor,"MYSQL","DATABASE", pcArcMySQL)

		luValor = "MySQL ODBC 3.51 Driver"
		WriteIniEntry(luValor,"MYSQL","DRIVER", pcArcMySQL)

		luValor = "localhost"
		WriteIniEntry(luValor,"MYSQL","SERVER", pcArcMySQL)

		luValor = "root"
		WriteIniEntry(luValor,"MYSQL","USER", pcArcMySQL)

		luValor = "1234"
		WriteIniEntry(luValor,"MYSQL","PASSWORD", pcArcMySQL)

	ELSE &&!FILE(pcArcMySQL)
	ENDIF &&!FILE(pcArcMySQL)

	gcPath = GetIniEntry("GENERAL", "CURDIR", pcArcMySQL)
	gcDirDat = GetIniEntry("GENERAL", "DIRDAT", pcArcMySQL)
	gcDirTempP = GetIniEntry("GENERAL", "DIRTEMP", pcArcMySQL)
	gcDirRep = GetIniEntry("GENERAL", "DIRREP", pcArcMySQL)
	gcDriver = GetIniEntry("MYSQL", "DRIVER", pcArcMySQL)
	gcPort = GetIniEntry("MYSQL", "PORT", pcArcMySQL)
	gcDataBase = GetIniEntry("MYSQL", "DATABASE", pcArcMySQL)
	gcServer = GetIniEntry("MYSQL", "SERVER", pcArcMySQL)
	gcUser = GetIniEntry("MYSQL", "USER", pcArcMySQL)
	gcPwd = GetIniEntry("MYSQL", "PASSWORD", pcArcMySQL)
ENDPROC

*//Click del CheckBox
*!*	DEFINE CLASS MYHANDLER AS SESSION
*!*		PROCEDURE CHK_CLICK
*!*		REPLACE MIGCON WITH PRINCIPAL.GRID1.COLUMN3.CHECK1.VALUE IN TmpNumTab
*!*		RETURN
*!*	ENDDEFINE

&&*-----------------------------------------------------------------------------------------------*
&&*... Prepara el formato de los datos para el INSERT o UPDATE en MySQL
&&*-----------------------------------------------------------------------------------------------*
FUNCTION DataFormat(tuDato)
	DO CASE
	CASE VARTYPE(tuDato) = "C"
		RETURN ("'"+ALLTRIM(STRTRAN(STRTRAN(tuDato,"'"," "),'"',' '))+"'")

	CASE VARTYPE(tuDato) = "N"
		RETURN (ALLTRIM(STR(tuDato,12,2)))

	CASE VARTYPE(tuDato) = "D"
		IF !EMPTY(tuDato)
			RETURN ("'"+PADL(ALLTRIM(STR(YEAR(tuDato))),4,'0')+"-"+PADL(ALLTRIM(STR(MONTH(tuDato))),2,'0')+"-"+PADL(ALLTRIM(STR(DAY(tuDato))),2,'0')+"'")
		ELSE &&!EMPTY(tuDato)
			RETURN ("'0000-00-00'")
		ENDIF &&!EMPTY(tuDato)
	CASE VARTYPE(tuDato) = "T"
		PRIVATE pcDateTime, pcHora
		pcDateTime = TTOC(tuDato)
		pcHora = RIGHT(pcDateTime,8)
		pcHora = STRTRAN(pcHora,'AM','')
		pcHora = STRTRAN(pcHora,'PM','')
		pcHora = STRTRAN(pcHora,'am','')
		pcHora = STRTRAN(pcHora,'pm','')
		pcHora = ALLTRIM(pcHora)
		tuDato = TTOD(tuDato)
		RETURN ("'"+PADL(ALLTRIM(STR(YEAR(tuDato))),4,'0')+"-"+PADL(ALLTRIM(STR(MONTH(tuDato))),2,'0')+"-"+PADL(ALLTRIM(STR(DAY(tuDato))),2,'0')+SPACE(1)+pcHora+"'")

	CASE VARTYPE(tuDato) = "U"
		RETURN ("0")

	CASE VARTYPE(tuDato) = "L"
		RETURN (IIF(tuDato,"'1'","'0'"))
	OTHERWISE
		RETURN ("0")
	ENDCASE

*----------------------------------------------------------------------------------------------*
*...FUNCTION WRITEINIENTRY
*...Escribe un valor de un archivo INI.
*...Si no existe el archivo, la sección o la entrada, la crea.
*...Retorna .T. si tuvo éxito
*...PARAMETROS:
*...tcFileName = Nombre y ruta completa del archivo.INI
*...tcSection = Sección del archivo.INI
*...tcEntry = Entrada del archivo.INI
*...tcValue = Valor de la entrada
*...USO: WriteFileIni("C:MiArchivo.ini","Default","Port","2")
*...RETORNO: Logico
*----------------------------------------------------
FUNCTION WriteIniEntry(tcValue,tcSection,tcEntry,tcFilename)
	DECLARE INTEGER WritePrivateProfileString ;
		IN WIN32API ;
		STRING cSection,STRING cEntry,STRING cEntry,;
		STRING cFileName

	RETURN IIF(WritePrivateProfileString(tcSection,tcEntry,tcValue,tcFilename)=1, .T., .F.)
ENDFUNC
*----------------------------------------------------------------------------------------------*
*...FUNCTION GETINIENTRY
*...Lee un valor de un archivo INI.
*...Si no existe el archivo, la sección o la entrada, retorna .NULL.
*...PARAMETROS:
*...tcFileName = Nombre y ruta completa del archivo.INI
*...tcSection = Sección del archivo.INI
*...tcEntry = Entrada del archivo.INI
*...USO: Getinientry("C:MiArchivo.ini","Default","Port")
*...RETORNO: Caracter
*----------------------------------------------------------------------------------------------*
FUNCTION GetIniEntry(tcSection,tcEntry,tcFilename)
	LOCAL LCINIVALUE, LNRESULT, LNBUFFERSIZE
	DECLARE INTEGER GetPrivateProfileString ;
		IN WIN32API ;
		STRING cSection,;
		STRING cEntry,;
		STRING cDefault,;
		STRING @cRetVal,;
		INTEGER nSize,;
		STRING cFileName
	LNBUFFERSIZE = 255
	LCINIVALUE = SPAC(LNBUFFERSIZE)
	LNRESULT=GetPrivateProfileString(tcSection,tcEntry,"*NULL*",;
		@LCINIVALUE,LNBUFFERSIZE,tcFilename)
	LCINIVALUE=SUBSTR(LCINIVALUE,1,LNRESULT)
	IF LCINIVALUE="*NULL*"
		LCINIVALUE=.NULL.
	ENDIF
	RETURN LCINIVALUE
ENDFUNC

FUNCTION ISRUNNING(tcProgram)
	DECLARE INTEGER GetActiveWindow 	IN Win32API
	DECLARE INTEGER GetWindow 			IN WIN32API INTEGER HWND, INTEGER nType
	DECLARE INTEGER GetWindowText 		IN Win32API INTEGER HWND, STRING @cText, INTEGER nType
	DECLARE INTEGER BringWindowToTop 	IN Win32API INTEGER HWND

	nHand 	 = CHECKRUNNING(tcProgram)
	IF nHand = 0
		RETURN .F.
	ELSE
		BringWindowToTop(nHand)
	ENDIF
ENDFUNC

FUNCTION CHECKRUNNING(cTitulo)
	hNext = GetActiveWindow()
	DO WHILE hNext<>0
		cText = REPLICATE(CHR(0),80)
		GetWindowText(hNext,@cText,80)
		IF UPPER(TRIM(cTitulo)) $ UPPER(cText)
			RETURN hNext
		ENDIF
		hNext = GetWindow(hNext,2)
	ENDDO
	RETURN 0
ENDFUNC