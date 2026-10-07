@echo off
REM ------------------------------------------------------------------
REM Liga o Painel de Montagem SoBot. Dois cliques neste arquivo.
REM A janela preta que abre E o servidor: deixe-a aberta enquanto usar o painel.
REM Para desligar, feche a janela.
REM ------------------------------------------------------------------

REM Usa UTF-8 para os acentos aparecerem certos na janela
chcp 65001 >nul

REM Vai para a pasta onde este .bat esta (assim funciona de qualquer lugar)
cd /d "%~dp0"

REM Se faltar alguma biblioteca do Python, instala (so na primeira vez)
python -c "import flask, openpyxl, win32com" 2>nul || python -m pip install flask openpyxl pywin32

REM Inicia o servidor; ele abre o painel no navegador sozinho
python server.py
pause
