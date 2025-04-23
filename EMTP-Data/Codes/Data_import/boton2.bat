@echo off
setlocal enabledelayedexpansion

echo Intentando ejecutar convertidor.py con diferentes versiones de Python...

:: Intenta con Python 3.13
echo Probando con Python 3.13...
py -3.13 convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con Python 3.13
    goto fin
) else (
    echo Python 3.13 no disponible o falló la ejecución, probando siguiente versión...
)

:: Intenta con Python 3.12
echo Probando con Python 3.12...
py -3.12 convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con Python 3.12
    goto fin
) else (
    echo Python 3.12 no disponible o falló la ejecución, probando siguiente versión...
)

:: Intenta con Python 3.11
echo Probando con Python 3.11...
py -3.11 convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con Python 3.11
    goto fin
) else (
    echo Python 3.11 no disponible o falló la ejecución, probando siguiente versión...
)

:: Intenta con Python 3.10
echo Probando con Python 3.10...
py -3.10 convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con Python 3.10
    goto fin
) else (
    echo Python 3.10 no disponible o falló la ejecución, probando siguiente versión...
)

:: Intenta con Python genérico
echo Probando con Python genérico...
python convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con Python genérico
    goto fin
) else (
    echo Python genérico no disponible o falló la ejecución, probando siguiente versión...
)

:: Intenta con py genérico (última opción)
echo Probando con py genérico...
py convertidor.py 2>nul
if !errorlevel! equ 0 (
    echo Éxito con py genérico
    goto fin
) else (
    echo No se pudo ejecutar el script con ninguna versión de Python disponible.
    echo Asegúrate de que Python esté instalado y que el script no tenga errores.
)

:fin
pause