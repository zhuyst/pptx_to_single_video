@echo off
chcp 65001 >nul
echo ==========================================
echo PPT转视频工具 - 编译为EXE文件
echo ==========================================
echo.

echo [1/4] 检查并安装PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo 安装PyInstaller失败，请检查网络连接或手动安装
        pause
        exit /b 1
    )
) else (
    echo PyInstaller已安装
)

echo.
echo [2/4] 检查依赖项...
pip install -r requirements.txt
if errorlevel 1 (
    echo 安装依赖项失败
    pause
    exit /b 1
)

echo.
echo [3/4] 开始编译GUI程序为EXE文件...
echo 编译选项:
echo   - 单文件模式 (--onefile)
echo   - 无控制台窗口 (--windowed)
echo   - 包含所有依赖项
echo.

pyinstaller --onefile --windowed --name="PPT转视频工具" --icon=NONE --add-data "requirements.txt;." main_gui.py

if errorlevel 1 (
    echo.
    echo 编译失败！可能的原因：
    echo 1. 缺少必要的依赖项
    echo 2. Python环境配置问题
    echo 3. 文件权限问题
    echo.
    pause
    exit /b 1
)

echo.
echo [4/4] 编译完成！
echo.
echo 输出位置: dist\PPT转视频工具.exe
echo.

if exist "dist\PPT转视频工具.exe" (
    echo ✓ EXE文件已成功创建
    echo 文件大小: 
    for %%F in ("dist\PPT转视频工具.exe") do echo   %%~zF 字节
    echo.
    echo 清理临时文件...
    if exist build rmdir /s /q build
    if exist "PPT转视频工具.spec" del "PPT转视频工具.spec"
    echo.
    echo 是否立即运行生成的EXE文件进行测试？
    set /p test="输入 y 运行测试，任意键跳过: "
    if /i "%test%"=="y" (
        echo 正在启动EXE文件...
        start "" "dist\PPT转视频工具.exe"
    )
) else (
    echo ✗ EXE文件创建失败
)

echo.
echo ==========================================
echo 编译过程完成
echo ==========================================
pause 