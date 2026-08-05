@echo off
Rem ***********************************************************************
Rem *                 Procedimiento de Descompresion 
Rem ***********************************************************************
Rem
Rem                     Depto. Canales At.Clientes
Rem
Rem Se realiza copia de los archivos instalables sin restriccion de fechas
Rem Se recomienda que tales archivos no sean ejecutables
Rem 

        BE ROWCOL 7,25 "        NORMALIZANDO ESTACION POSVENTA"

C:\BIN\F\be window 0,0,24,79 white on blue
C:\BIN\F\be window 8,10,16,70 black on cyan explode shadow
C:\BIN\F\be rowcol 9,18  "   NORMALIZANDO DIRECTORIO    "
C:\BIN\F\be rowcol 12,23 "        ESPERE UN MOMENTO ..."

XCOPY C:\BCONVSDW C:\ /S /E /V /Y >nul
DELTREE /Y C:\BCONVSDW         >nul

Rem ** EN ESTE ESPACIO PUEDE INCLUIR NORMALIZACION PROPIA DEL PRODUCTO

IF EXIST C:\BIN\S\EXECTL00.EXE DEL C:\BIN\S\EXECTL00.EXE >NUL

Rem ** FIN DEL ESPACIO DE NORMALIZACION PROPIA

:OK
C:\BIN\F\be rowcol 14,23 "        PROCESO TERMINADO   " 
