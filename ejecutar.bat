@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Instalando dependencias (solo la primera vez tarda)...
python -m pip install -q -r requirements.txt || py -m pip install -q -r requirements.txt
python reformateador.py || py reformateador.py
pause
