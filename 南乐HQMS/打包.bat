@echo off
chcp 65001 > nul
echo ========================================
echo   南乐HQMS上报系统 - 打包程序
echo ========================================
echo.

echo 正在安装依赖...
pip install pyinstaller flask pyodbc pandas -q

echo.
echo 正在打包，请稍候...
pyinstaller 南乐hqms上报.spec --clean

echo.
echo ========================================
echo   打包完成！
echo ========================================
echo.
echo 可执行文件位于: dist\南乐HQMS上报系统.exe
echo.
echo 启动程序请双击: dist\南乐HQMS上报系统.exe
echo.
pause
