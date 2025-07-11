@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ==========================================
echo PPT转视频工具 - 高级EXE编译脚本
echo ==========================================
echo.

echo [信息] 开始编译流程...
echo 当前目录: %CD%
echo Python版本:
python --version
echo.

echo [1/6] 检查必要文件...
if not exist "main_gui.py" (
    echo ✗ 错误: 找不到main_gui.py文件
    pause
    exit /b 1
)
if not exist "requirements.txt" (
    echo ✗ 错误: 找不到requirements.txt文件
    pause
    exit /b 1
)
echo ✓ 必要文件检查完成

echo.
echo [2/6] 安装和升级构建工具...
pip install --upgrade pip setuptools wheel
pip install --upgrade pyinstaller>=5.0

echo.
echo [3/6] 安装项目依赖...
pip install -r requirements.txt

echo.
echo [4/6] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "*.spec" del "*.spec"
echo ✓ 清理完成

echo.
echo [5/6] 开始编译EXE文件...
echo 编译配置:
echo   - 单文件打包 (--onefile)
echo   - 窗口模式，无控制台 (--windowed)
echo   - 优化导入 (--optimize=2)
echo   - 包含隐藏导入
echo   - 自定义名称和图标
echo.

:: 创建临时的PyInstaller配置
echo 正在配置PyInstaller...

pyinstaller ^
    --onefile ^
    --windowed ^
    --optimize=2 ^
    --name="PPT转视频工具" ^
    --distpath="dist" ^
    --workpath="build" ^
    --specpath="." ^
    --hidden-import=win32com.client ^
    --hidden-import=tkinter ^
    --hidden-import=tkinterdnd2 ^
    --hidden-import=threading ^
    --hidden-import=subprocess ^
    --add-data="requirements.txt;." ^
    --collect-submodules=tkinterdnd2 ^
    --collect-submodules=win32com ^
    --noupx ^
    --console=False ^
    main_gui.py

if errorlevel 1 (
    echo.
    echo ✗ 编译失败！
    echo.
    echo 常见解决方案:
    echo 1. 确保所有依赖项已正确安装
    echo 2. 尝试在虚拟环境中编译
    echo 3. 检查Python和pip版本
    echo 4. 手动安装缺失的模块
    echo.
    echo 详细错误信息请查看上方输出
    pause
    exit /b 1
)

echo.
echo [6/6] 验证编译结果...

if exist "dist\PPT转视频工具.exe" (
    echo ✓ EXE文件编译成功！
    echo.
    echo 文件信息:
    echo   位置: %CD%\dist\PPT转视频工具.exe
    for %%F in ("dist\PPT转视频工具.exe") do (
        echo   大小: %%~zF 字节 ^(%.1f MB^)
        set /a size_mb=%%~zF/1024/1024
        echo   大小: !size_mb! MB
    )
    echo   创建时间: 
    forfiles /p "dist" /m "PPT转视频工具.exe" /c "cmd /c echo @fdate @ftime"
    
    echo.
    echo [后续操作]
    echo ✓ 清理构建临时文件...
    if exist build rmdir /s /q build
    if exist "PPT转视频工具.spec" del "PPT转视频工具.spec"
    
    echo.
    echo [测试选项]
    set /p test="是否立即测试EXE文件？(y/n): "
    if /i "!test!"=="y" (
        echo 正在启动EXE文件进行测试...
        start "" "dist\PPT转视频工具.exe"
        echo 请检查程序是否正常启动和运行
    )
    
    echo.
    echo [分发说明]
    echo 要分发此程序，请确保目标计算机具备:
    echo 1. Windows 10 或更高版本
    echo 2. Microsoft PowerPoint 2016 或更高版本
    echo 3. Visual C++ Redistributable
    echo.
    echo EXE文件位置: dist\PPT转视频工具.exe
    echo 此文件可以独立运行，无需安装Python环境
    
) else (
    echo ✗ EXE文件创建失败
    echo 请检查上方的错误信息
)

echo.
echo ==========================================
echo 编译流程完成
echo ==========================================
echo 按任意键退出...
pause >nul 