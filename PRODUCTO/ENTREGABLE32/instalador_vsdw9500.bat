@echo off
setlocal EnableExtensions EnableDelayedExpansion
rem ===========================================================================
rem  VSDW9500 - Visado de Productos (VB6 32 bits)
rem  Instalador de estacion. Ver docs/empaquetado-y-despliegue.md del framework.
rem
rem  Uso:
rem    instalador_vsdw9500.bat [/SOURCE ruta] [/SRMHOST host] [/SRMPORT puerto]
rem                            [/CHECKONLY]
rem
rem  /SRMHOST solo se usa si la estacion NO tiene C:\Windows\Srmw.ini.
rem  Si ya existe, NO se toca: es un archivo compartido con otros
rem  aplicativos del banco.
rem
rem  Sin acentos a proposito: la consola en CP850 los muestra mal.
rem ===========================================================================

set "SRC=%~dp0"
if "%SRC:~-1%"=="\" set "SRC=%SRC:~0,-1%"
set "SRMHOST="
set "SRMPORT=6736"
set "CHECKONLY=0"
set "VERIFY_FAIL=0"

rem --- ARG: acepta la forma "/X valor" y la forma "/X=valor" -----------------
:ARG
if "%~1"=="" goto ARG_FIN
set "A=%~1"
if /I "%A%"=="/CHECKONLY" (set "CHECKONLY=1" & shift & goto ARG)
if /I "%A%"=="/SOURCE"    (set "SRC=%~2"     & shift & shift & goto ARG)
if /I "%A%"=="/SRMHOST"   (set "SRMHOST=%~2" & shift & shift & goto ARG)
if /I "%A%"=="/SRMPORT"   (set "SRMPORT=%~2" & shift & shift & goto ARG)
for /f "tokens=1,* delims==" %%a in ("%A%") do (
  if /I "%%a"=="/SOURCE"  set "SRC=%%b"
  if /I "%%a"=="/SRMHOST" set "SRMHOST=%%b"
  if /I "%%a"=="/SRMPORT" set "SRMPORT=%%b"
)
shift
goto ARG
:ARG_FIN

echo.
echo ===========================================================
echo  VSDW9500 - Visado de Productos (32 bits)
echo ===========================================================
echo  Origen: %SRC%
if "%CHECKONLY%"=="1" echo  Modo   : SOLO VERIFICACION (no copia nada)
echo.

if "%CHECKONLY%"=="1" goto VERIFICACION

rem --- LOCATE ----------------------------------------------------------------
if not exist "%SRC%\BIN\S\VSDPROD.EXE" (
  echo [ERROR] No se encontro %SRC%\BIN\S\VSDPROD.EXE
  echo         Indique la ruta del paquete con /SOURCE=ruta
  goto FIN_ERROR
)

rem --- ADMIN -----------------------------------------------------------------
net session >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Ejecute este instalador como Administrador.
  goto FIN_ERROR
)

rem --- DIRS ------------------------------------------------------------------
echo [1/7] Creando directorios...
if not exist "C:\Bin\S"        md "C:\Bin\S"
if not exist "C:\Bin\F"        md "C:\Bin\F"
if not exist "C:\Data\Visado"  md "C:\Data\Visado"
if not exist "C:\Etc\Bin"      md "C:\Etc\Bin"

rem --- PRESERVE: la config de la estacion no se pisa en un upgrade -----------
echo [2/7] Preservando configuracion de la estacion...
if exist "C:\Data\Visado\Visado.INI" (
  copy /Y "C:\Data\Visado\Visado.INI" "%TEMP%\Visado.INI.bak" >nul
  echo       Visado.INI existente respaldado
  set "TENIA_INI=1"
) else (
  set "TENIA_INI=0"
)

rem --- COPYTREE: xcopy, NO robocopy (no existe en XP) ------------------------
echo [3/7] Copiando archivos...
xcopy "%SRC%\BIN\S\*"       "C:\Bin\S\"       /E /I /Y /Q >nul
xcopy "%SRC%\BIN\F\*"       "C:\Bin\F\"       /E /I /Y /Q >nul
xcopy "%SRC%\DATA\VISADO\*" "C:\Data\Visado\" /E /I /Y /Q >nul
xcopy "%SRC%\ETC\BIN\*"     "C:\Etc\Bin\"     /E /I /Y /Q >nul

rem --- RESTORE ---------------------------------------------------------------
if "!TENIA_INI!"=="1" (
  copy /Y "%TEMP%\Visado.INI.bak" "C:\Data\Visado\Visado.INI" >nul
  del /q "%TEMP%\Visado.INI.bak" >nul 2>&1
  echo [4/7] Configuracion de la estacion restaurada
) else (
  echo [4/7] Primera instalacion: se deja el Visado.INI del paquete
)

rem --- REGISTER_OCX ----------------------------------------------------------
echo [5/7] Registrando controles...
rem  En un Windows de 64 bits hay que registrar los OCX de 32 bits con el
rem  regsvr32 de SysWOW64; el de system32 es el de 64 bits y falla.
set "REGSVR=regsvr32"
if exist "%SystemRoot%\SysWOW64\regsvr32.exe" set "REGSVR=%SystemRoot%\SysWOW64\regsvr32.exe"
for %%O in (THREED32.OCX MSFLXGRD.OCX COMDLG32.OCX MSCOMCT2.OCX) do (
  if exist "C:\Etc\Bin\%%O" "!REGSVR!" /s "C:\Etc\Bin\%%O"
)
if exist "C:\Etc\Bin\licencias.reg" regedit /s "C:\Etc\Bin\licencias.reg"

rem --- SRM -------------------------------------------------------------------
rem  C:\Windows\Srmw.ini es un archivo DELICADO de la estacion: lo comparten
rem  otros aplicativos del banco y tiene la configuracion real de nodos. Este
rem  instalador NO lo pisa NUNCA si ya existe, aunque se pase /SRMHOST.
rem  Solo lo crea cuando no existe (estacion nueva), y ahi si hace falta el
rem  /SRMHOST del ambiente destino.
if exist "C:\Windows\Srmw.ini" (
  echo [6/7] C:\Windows\Srmw.ini ya existe: NO se toca ^(archivo compartido^)
  for /f "tokens=2 delims==" %%h in ('findstr /B /I "Host=" "C:\Windows\Srmw.ini"') do echo       Host actual = %%h
) else (
  if not "%SRMHOST%"=="" (
    echo [6/7] No existe Srmw.ini: se crea con Host=%SRMHOST%
    (echo [SRM])> "C:\Windows\Srmw.ini"
    (echo Host=%SRMHOST%)>> "C:\Windows\Srmw.ini"
    (echo TCP_Port_Srm=%SRMPORT%)>> "C:\Windows\Srmw.ini"
    (echo.)>> "C:\Windows\Srmw.ini"
    (echo [TCP])>> "C:\Windows\Srmw.ini"
    (echo TCP_Port_Srm=%SRMPORT%)>> "C:\Windows\Srmw.ini"
  ) else (
    echo [6/7] FALTA C:\Windows\Srmw.ini y no se indico /SRMHOST
    echo       Pidalo al area de infraestructura o corra con /SRMHOST host
  )
)

rem --- OFICINA ---------------------------------------------------------------
rem  La variable OFICINA la asigna el SOFTWARE BASICO de la estacion, que no
rem  es parte de este entregable. Este instalador NO la escribe: solo informa
rem  si esta puesta. Si falta, hay que pedirla al area que administra la
rem  estacion -- ponerla a mano desde aca podria dejarla en desacuerdo con la
rem  oficina real y hacer que las consultas salgan mal.
if not "%OFICINA%"=="" (
  echo [7/7] OFICINA=%OFICINA% ^(la asigna el software basico, no el instalador^)
) else (
  echo [7/7] OFICINA no esta definida en esta estacion
  echo       La asigna el software basico. Verifique con el area que administra
  echo       la estacion antes de usar el aplicativo.
)

rem ===========================================================================
:VERIFICACION
echo.
echo ===========================================================
echo  VERIFICACION
echo ===========================================================

call :EXISTS "VSDPROD.EXE"  "C:\Bin\S\VSDPROD.EXE"
call :EXISTS "SRMW32.DLL"   "C:\Bin\F\SRMW32.DLL"
call :EXISTS "Visado.MDB"   "C:\Data\Visado\Visado.MDB"
call :EXISTS "EnvioSRM.MDB" "C:\Data\Visado\EnvioSRM.MDB"
call :EXISTS "Oficina.MDB"  "C:\Data\Visado\Oficina.MDB"
call :EXISTS "Notarios.MDB" "C:\Data\Visado\Notarios.MDB"
call :EXISTS "Visado.INI"   "C:\Data\Visado\Visado.INI"

echo  --- prerrequisitos de ambiente (no vienen en el paquete) ---
rem  El aplicativo es de 32 bits. En un Windows de 64 bits (Win7/10/11) su
rem  runtime NO esta en system32 sino en SysWOW64, y DAO cuelga de
rem  "Program Files (x86)". Verificado en Windows 11 24H2: ambos vienen ya
rem  instalados de fabrica, no hay que agregarlos.
if exist "%SystemRoot%\SysWOW64\MSVBVM60.DLL" goto :PRE_VB6_OK
call :EXISTS "MSVBVM60.DLL (runtime VB6)" "%SystemRoot%\system32\MSVBVM60.DLL"
goto :PRE_DAO
:PRE_VB6_OK
echo   [OK]    MSVBVM60.DLL (runtime VB6, SysWOW64)
:PRE_DAO
set "DAOX86=%SystemDrive%\Program Files (x86)\Common Files\Microsoft Shared\DAO\DAO360.DLL"
if exist "!DAOX86!" goto :PRE_DAO_OK
call :EXISTS "DAO360.DLL   (DAO 3.6)" "%SystemDrive%\Program Files\Common Files\Microsoft Shared\DAO\DAO360.DLL"
goto :PRE_FIN
:PRE_DAO_OK
echo   [OK]    DAO360.DLL   (DAO 3.6, Program Files x86)
:PRE_FIN

echo  --- controles registrados (por TypeLib, no por CLSID) ---
call :OCXCHECK "THREED32" "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}"
call :OCXCHECK "MSFLXGRD" "{5E9E78A0-531B-11CF-91F6-C2863C385E30}"
call :OCXCHECK "COMDLG32" "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}"

echo  --- oficina (la asigna el software basico de la estacion) ---
if not "%OFICINA%"=="" (echo   [OK]    OFICINA=%OFICINA%) else (echo   [AVISO] OFICINA sin definir -- reclamar al administrador de la estacion)
echo  --- SRM ---
call :SRMCHECK

echo.
if "!VERIFY_FAIL!"=="1" (
  echo  RESULTADO: HAY FALTANTES -- revise las lineas [FALTA] de arriba.
  goto FIN_ERROR
)
echo  RESULTADO: OK
echo.
goto FIN_OK

rem ===========================================================================
rem  Subrutinas
rem ===========================================================================
:EXISTS
rem  Sin bloques ( ... ): %~1 quita las comillas, asi que una ruta como
rem  "C:\Program Files (x86)\..." mete un ')' que cierra el bloque y cmd
rem  aborta con "No se esperaba ... en este momento". En XP no se ve porque
rem  ahi la ruta no tiene parentesis; en Windows 11 x64 revienta.
if exist "%~2" goto :EXISTS_OK
echo   [FALTA] %~1 -- %~2
set "VERIFY_FAIL=1"
exit /b 0
:EXISTS_OK
echo   [OK]    %~1
exit /b 0

:OCXCHECK
rem El GUID que declara el .VBP es el de la TypeLib, NO el CLSID del coclass:
rem verificar contra CLSID da un falso "no registrado" aunque el control ande.
rem  En x64 los OCX de 32 bits quedan bajo Wow6432Node: se miran las dos vistas.
reg query "HKLM\SOFTWARE\Classes\TypeLib\%~2" >nul 2>&1
if not errorlevel 1 goto :OCXCHECK_OK
reg query "HKLM\SOFTWARE\Classes\Wow6432Node\TypeLib\%~2" >nul 2>&1
if not errorlevel 1 goto :OCXCHECK_OK
echo   [FALTA] OCX %~1 no registrado
set "VERIFY_FAIL=1"
exit /b 0
:OCXCHECK_OK
echo   [OK]    OCX %~1 registrado
exit /b 0

:SRMCHECK
if not exist "C:\Windows\Srmw.ini" (
  echo   [FALTA] C:\Windows\Srmw.ini -- use /SRMHOST para generarlo
  set "VERIFY_FAIL=1"
  exit /b 0
)
findstr /B /I "Host=" "C:\Windows\Srmw.ini" >nul 2>&1
if errorlevel 1 (echo   [FALTA] Srmw.ini sin Host=& set "VERIFY_FAIL=1") else (
  for /f "tokens=2 delims==" %%h in ('findstr /B /I "Host=" "C:\Windows\Srmw.ini"') do echo   [OK]    SRM Host=%%h
)
exit /b 0

:FIN_ERROR
echo.
endlocal
exit /b 1

:FIN_OK
endlocal
exit /b 0
