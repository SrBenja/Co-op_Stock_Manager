@echo off
REM ======================================================================
REM 0) LIMPIAR carpetas anteriores: build\ y dist\
REM ======================================================================
if exist build rd /s /q build
if exist dist  rd /s /q dist

REM ======================================================================
REM 1) (Opcional) Asegurarse de tener pip al día
REM ======================================================================
echo Instalando/actualizando pip...
python -m pip install --upgrade pip

REM ======================================================================
REM 2) Empaquetar con PyInstaller en un único exe GUI dentro de dist\
REM ======================================================================
echo Empaquetando con PyInstaller...
pyinstaller --noconfirm --clean --onefile --noupx ^
            --distpath dist ^
            --workpath build ^
            --specpath build ^
            --name Co-op_Stock_Manager ^
            --windowed ^
            main.py

REM ======================================================================
REM 3) Eliminar carpeta build\ generada por PyInstaller
REM ======================================================================
if exist build rd /s /q build

REM ======================================================================
REM 4) Crear dentro de dist\ las carpetas que tu app necesita
REM ======================================================================
if not exist "dist\backup" mkdir "dist\backup"
if not exist "dist\config" mkdir "dist\config"

REM ======================================================================
REM 5) FIN — ahora ejecuta tu .exe manualmente cuando lo desees
REM ======================================================================
exit /b