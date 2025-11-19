@echo off
rem Wrapper interactivo para setDate
setlocal
cd /d "%~dp0"
echo.
echo === Ejecutar setDate (wrapper interactivo) ===
echo.
set /p FILE=Nombre o ruta del archivo (ej: Libro1.xlsx): 
if "%FILE%"=="" (
  echo Archivo no proporcionado. Saliendo.
  pause
  exit /b 1
)
set /p DATE=Fecha en formato YYYYMMDDhhmm (ej: 200901021200): 
if "%DATE%"=="" (
  echo Fecha no proporcionada. Saliendo.
  pause
  exit /b 1
)
echo.
echo Ejecutando: setDate.bat "%FILE%" %DATE%
call "%~dp0setDate.bat" "%FILE%" %DATE%
echo.
echo Proceso finalizado.
pause
