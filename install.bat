@echo off
chcp 65001 >nul
echo 正在安装依赖...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 安装失败，请检查Python环境
    pause
    exit /b 1
)
echo 依赖安装完成！
pause
