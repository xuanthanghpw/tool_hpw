@echo off
echo =========================================
echo   TU DONG DONG GOI SMARTY TOOL
echo =========================================

echo 1. Cai dat PyInstaller (neu chua co)...
pip install pyinstaller

echo.
echo 2. Dang tien hanh Build file EXE (Vui long cho vai phut)...
pyinstaller --noconsole --onefile smarty.py

echo.
echo 3. Dang tao thu muc dong goi (SmartyTool_Release)...
if exist "SmartyTool_Release" rmdir /s /q "SmartyTool_Release"
mkdir SmartyTool_Release
mkdir SmartyTool_Release\data

echo.
echo 4. Dang di chuyen file exe va copy data...
move dist\smarty.exe SmartyTool_Release\ >nul
xcopy data\* SmartyTool_Release\data\ /E /H /C /I /Y >nul

echo.
echo 5. Don dep cac file rac sinh ra trong luc build...
rmdir /s /q build
rmdir /s /q dist
if exist smarty.spec del smarty.spec

echo.
echo =========================================
echo HOAN TAT!
echo Toan bo san pham da nam gon trong thu muc: SmartyTool_Release
echo Ban co the nen thu muc nay lai (ZIP/RAR) va mang di bat cu dau de dung.
echo =========================================
pause