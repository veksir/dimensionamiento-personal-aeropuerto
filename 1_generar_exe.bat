@echo off
REM ============================================================
REM  Genera CalculadoraPersonal.exe a partir de CalculadoraPersonal.py
REM  Esto se ejecuta UNA SOLA VEZ, en un computador que si tenga
REM  Python instalado (por ejemplo el tuyo). El resultado (el .exe)
REM  es el que le entregas al usuario final, que no necesita
REM  instalar nada.
REM ============================================================

echo Instalando dependencias necesarias...
pip install pandas numpy openpyxl plotly tkinterdnd2 pyinstaller

echo.
echo Generando el programa (carpeta, no archivo unico -- asi abre
echo de inmediato en vez de tardar unos segundos cada vez que se abre)...
pyinstaller --onedir --windowed --name "CalculadoraPersonal" --collect-all tkinterdnd2 CalculadoraPersonal.py

echo.
echo ============================================================
echo  Listo. Se creo la carpeta:
echo    dist\CalculadoraPersonal
echo  Esa carpeta COMPLETA es lo que le compartes al usuario final
echo  (comprimida en un .zip, por ejemplo). Dentro esta
echo  CalculadoraPersonal.exe, que es el que debe abrir con doble clic.
echo  No se puede compartir solo el .exe suelto: necesita los demas
echo  archivos de la carpeta para funcionar.
echo ============================================================
pause
