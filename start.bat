@echo off
chcp 65001 >nul
setlocal

:: 获取脚本所在目录并切换到该目录
cd /d "%~dp0"

echo 🚀 启动图片转Word工具...
echo 📁 当前目录: %CD%
echo.

:: 检查Node.js是否安装
where node >nul 2>&1
if errorlevel 1 (
    echo ❌ Node.js 未安装，请先安装 Node.js
    pause
    exit /b 1
)

:: 运行程序
echo 🎯 开始处理图片...
call npm start

echo.
echo ✅ 处理完成！
echo 📄 生成的Word文档: output.docx
echo.

:: 打开生成的Word文档
if exist "output.docx" (
    start "" "output.docx"
) else (
    echo ⚠️  未找到生成的Word文档
)

echo.
echo 🔄 3秒后自动关闭窗口...
timeout /t 3 /nobreak >nul

:: 自动关闭当前命令提示符窗口
exit 