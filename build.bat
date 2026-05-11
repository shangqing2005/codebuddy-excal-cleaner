@echo off
echo 开始打包 Excel数据清洗工具 v5.0
echo.

:: 检查是否有Python和PyInstaller
python --version
if errorlevel 1 (
    echo 错误：未找到Python，请确保Python已安装并添加到PATH
    pause
    exit /b 1
)

:: 确保xlsx_shifter.exe存在
if not exist "xlsx_shifter.exe" (
    echo 错误：未找到xlsx_shifter.exe，请先编译Go引擎
    echo cd xlsx_shifter
    echo go build -o ../xlsx_shifter.exe main.go
    pause
    exit /b 1
)

:: 使用PyInstaller打包
echo 运行PyInstaller...
pyinstaller --onefile --windowed --icon="数据清洗.ico" --add-data="xlsx_shifter.exe;." --name="Excel数据清洗工具_v5.0" excel_cleaner_1.0.py

if errorlevel 1 (
    echo.
    echo 打包失败！请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包成功！
echo 输出文件: dist\Excel数据清洗工具_v5.0.exe
echo ========================================
echo.

:: 如果有之前的压缩包，移动到旧版本目录
if exist "Excel数据清洗工具*.zip" (
    echo 移动旧版本压缩包...
    if not exist "以前的压缩包" mkdir "以前的压缩包"
    move "Excel数据清洗工具*.zip" "以前的压缩包\" 2>nul
)

echo.
echo 打包完成！按任意键退出...
pause >nul
