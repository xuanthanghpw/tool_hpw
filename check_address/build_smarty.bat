@echo off
setlocal

echo =========================================
echo   TU DONG DONG GOI SMARTY TOOL
echo =========================================

echo.
echo 1. Kiem tra PyInstaller...
python -m PyInstaller --version

if errorlevel 1 (
    echo.
    echo [LOI] Khong tim thay PyInstaller.
    echo Dang cai dat PyInstaller...
    python -m pip install pyinstaller

    if errorlevel 1 (
        echo.
        echo [LOI] Cai dat PyInstaller that bai!
        pause
        exit /b 1
    )
)

echo.
echo 2. Dang tien hanh Build file EXE...
echo Vui long cho vai phut...
echo.

python -m PyInstaller --noconsole --onefile smarty.py

if errorlevel 1 (
    echo.
    echo =========================================
    echo [LOI] BUILD THAT BAI!
    echo =========================================
    echo Kiem tra loi phia tren de biet nguyen nhan.
    pause
    exit /b 1
)

if not exist "dist\smarty.exe" (
    echo.
    echo [LOI] Khong tim thay dist\smarty.exe sau khi build!
    pause
    exit /b 1
)

echo.
echo 3. Dang tao thu muc dong goi (SmartyTool_Release)...

if exist "SmartyTool_Release" (
    rmdir /s /q "SmartyTool_Release"
)

mkdir "SmartyTool_Release"

if exist "data" (
    mkdir "SmartyTool_Release\data"
)

echo.
echo 4. Dang di chuyen file EXE...

move /Y "dist\smarty.exe" "SmartyTool_Release\smarty.exe"

if errorlevel 1 (
    echo.
    echo [LOI] Khong the di chuyen file EXE!
    pause
    exit /b 1
)

echo.
echo 5. Dang copy data...

if exist "data" (
    xcopy "data\*" "SmartyTool_Release\data\" /E /H /C /I /Y
)

echo.
echo 6. Don dep cac file rac sinh ra trong luc build...

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "smarty.spec" del /q "smarty.spec"

echo.
echo =========================================
echo HOAN TAT!
echo =========================================
echo.
echo File EXE:
echo SmartyTool_Release\smarty.exe
echo.
echo Toan bo san pham da nam gon trong thu muc:
echo SmartyTool_Release
echo.
echo Ban co the nen thu muc nay thanh ZIP/RAR.
echo =========================================

pause