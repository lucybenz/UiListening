@echo off
chcp 65001 >nul
echo 正在安装依赖...
pip install -r requirements.txt
echo.
echo 启动程序...
python main.py
pause
