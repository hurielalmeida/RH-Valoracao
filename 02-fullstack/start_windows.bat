@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo [1/2] Criando ambiente virtual...
  py -3 -m venv .venv
  if errorlevel 1 (
    echo Nao foi possivel criar o ambiente virtual. Verifique se o Python 3 esta instalado.
    pause
    exit /b 1
  )
  echo [2/2] Instalando dependencias...
  .venv\Scripts\python.exe -m pip install --upgrade pip
  .venv\Scripts\python.exe -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Falha na instalacao das dependencias.
    pause
    exit /b 1
  )
) else (
  echo Ambiente virtual encontrado.
)
echo.
echo Iniciando RH Valoracao Full Stack V4.1...
echo Abra no navegador: http://127.0.0.1:8000
echo.
.venv\Scripts\python.exe -m uvicorn backend.app:app --host 127.0.0.1 --port 8000
pause
