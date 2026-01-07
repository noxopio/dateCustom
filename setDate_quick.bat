@echo off
rem Wrapper interactivo para setDate con soporte multi-archivo y modo continuo
setlocal EnableDelayedExpansion
cd /d "%~dp0"

:MENU
cls
echo.
echo ====================================================
echo   SETDATE - Modificador de fechas de archivos
echo ====================================================
echo.
echo Instrucciones:
echo   - Ingrese archivos y fechas en pares
echo   - Formato: archivo1.txt fecha1 archivo2.txt fecha2 ...
echo   - Fecha formato: YYYYMMDDhhmm (ej: 200901021200)
echo   - Escriba 'salir' para terminar
echo.
echo ====================================================
echo.

set /p INPUT=Ingrese archivos y fechas (o 'salir'): 

if /i "!INPUT!"=="salir" (
  echo.
  echo Saliendo...
  timeout /t 2 /nobreak >nul
  exit /b 0
)

if /i "!INPUT!"=="exit" (
  echo.
  echo Saliendo...
  timeout /t 2 /nobreak >nul
  exit /b 0
)

if "!INPUT!"=="" (
  echo.
  echo Error: No se proporcionaron datos.
  echo.
  pause
  goto MENU
)

echo.
echo Procesando archivos...
echo.

rem Llamar al script PowerShell con todos los parámetros
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setDate.ps1" !INPUT!

echo.
echo ====================================================
echo Proceso completado. Presione cualquier tecla para continuar...
echo ====================================================
pause >nul

goto MENU
