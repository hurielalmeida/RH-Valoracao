@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ============================================================
REM RH VALORACAO - GERADOR DO APLICATIVO WINDOWS V5.6
REM Build robusto em staging curto e isolado.
REM ============================================================

set "PROJECT_DIR=%~dp0"
set "STAGE=%TEMP%\RV55_BUILD"
set "STAGE_SRC=%STAGE%\src"
set "STAGE_VENV=%STAGE%\venv"
set "STAGE_BUILD=%STAGE_SRC%\build"
set "STAGE_DIST=%STAGE_SRC%\dist"
set "FINAL_DIST=%PROJECT_DIR%dist"
set "STAGE_EXE=%STAGE_DIST%\RH_Valoracao.exe"
set "FINAL_EXE=%FINAL_DIST%\RH_Valoracao.exe"

cd /d "%PROJECT_DIR%"

if not exist "%PROJECT_DIR%RH_Valoracao.spec" goto erro_estrutura
if not exist "%PROJECT_DIR%desktop_launcher.py" goto erro_estrutura
if not exist "%PROJECT_DIR%backend\app.py" goto erro_estrutura
if not exist "%PROJECT_DIR%core\rh_valoracao_core.py" goto erro_estrutura
if not exist "%PROJECT_DIR%frontend\index.html" goto erro_estrutura

cls
echo ============================================================
echo RH VALORACAO - GERADOR DO APLICATIVO WINDOWS V5.6
echo ============================================================
echo.
echo [1/6] Verificando Python...
where py >nul 2>nul
if errorlevel 1 goto erro_python
py -3 --version
if errorlevel 1 goto erro_python

 echo.
echo [2/6] Limpando area de build anterior...
if exist "%STAGE%" rmdir /s /q "%STAGE%" >nul 2>nul
if exist "%STAGE%" goto erro_limpeza
mkdir "%STAGE_SRC%" >nul 2>nul
if not exist "%STAGE_SRC%" goto erro_stage

 echo.
echo [3/6] Preparando copia minima do projeto...
robocopy "%PROJECT_DIR%backend" "%STAGE_SRC%\backend" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto erro_copia
robocopy "%PROJECT_DIR%core" "%STAGE_SRC%\core" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto erro_copia
robocopy "%PROJECT_DIR%frontend" "%STAGE_SRC%\frontend" /E /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto erro_copia

for %%F in (desktop_launcher.py RH_Valoracao.spec requirements.txt requirements-build.txt) do (
    if not exist "%PROJECT_DIR%%%F" (
        echo ERRO: arquivo obrigatorio ausente: %%F
        goto erro_copia
    )
    copy /y "%PROJECT_DIR%%%F" "%STAGE_SRC%\%%F" >nul
    if errorlevel 1 goto erro_copia
)

if not exist "%STAGE_SRC%\RH_Valoracao.spec" goto erro_copia

 echo Criando ambiente isolado em caminho curto...
py -3 -m venv "%STAGE_VENV%"
if errorlevel 1 goto erro_venv

"%STAGE_VENV%\Scripts\python.exe" -m pip install --disable-pip-version-check --no-cache-dir -r "%STAGE_SRC%\requirements-build.txt"
if errorlevel 1 goto erro_dependencias

 echo.
echo [4/6] Gerando executavel...
cd /d "%STAGE_SRC%"
"%STAGE_VENV%\Scripts\python.exe" -m PyInstaller --noconfirm --clean --workpath "%STAGE_BUILD%" --distpath "%STAGE_DIST%" "%STAGE_SRC%\RH_Valoracao.spec"
if errorlevel 1 goto erro_pyinstaller

REM IMPORTANTE: o .spec atual gera um EXE one-file diretamente em dist\RH_Valoracao.exe.
REM A V5.4 procurava incorretamente dist\RH_Valoracao\RH_Valoracao.exe.
if not exist "%STAGE_EXE%" goto erro_exe

 echo.
echo [5/6] Copiando o executavel para a pasta final...
cd /d "%PROJECT_DIR%"
if exist "%FINAL_DIST%" rmdir /s /q "%FINAL_DIST%" >nul 2>nul
if exist "%FINAL_DIST%" goto erro_final_dist
mkdir "%FINAL_DIST%" >nul 2>nul
if not exist "%FINAL_DIST%" goto erro_final_dist

copy /y "%STAGE_EXE%" "%FINAL_EXE%" >nul
if errorlevel 1 goto erro_final_dist

if exist "%PROJECT_DIR%README_EXECUTAVEL.txt" copy /y "%PROJECT_DIR%README_EXECUTAVEL.txt" "%FINAL_DIST%\README_EXECUTAVEL.txt" >nul

if not exist "%FINAL_EXE%" goto erro_exe_final

 echo.
echo [6/6] Validando resultado e limpando staging...
for %%F in ("%FINAL_EXE%") do echo EXE criado: %%~zF bytes
rmdir /s /q "%STAGE%" >nul 2>nul

 echo.
echo ============================================================
echo CONCLUIDO COM SUCESSO!
echo ============================================================
echo.
echo Aplicativo:
echo %FINAL_EXE%
echo.
echo O aplicativo final nao depende do Python instalado no PC do usuario.
echo Para compartilhar, compacte o arquivo RH_Valoracao.exe.
echo.
pause
exit /b 0

:erro_estrutura
echo Estrutura do projeto incompleta. Verifique os arquivos obrigatorios.
goto erro
:erro_python
echo Python 3.12 nao foi encontrado.
goto erro
:erro_limpeza
echo Nao foi possivel limpar a area temporaria anterior: %STAGE%
goto erro
:erro_stage
echo Nao foi possivel criar a area de staging: %STAGE%
goto erro
:erro_copia
echo Falha ao preparar a copia minima do projeto para o build.
goto erro
:erro_venv
echo Falha ao criar o ambiente de build.
goto erro
:erro_dependencias
echo Falha ao instalar as dependencias do build.
goto erro
:erro_pyinstaller
echo O PyInstaller falhou durante a geracao do executavel.
goto erro
:erro_exe
echo O PyInstaller terminou, mas o RH_Valoracao.exe nao foi encontrado em:
echo %STAGE_EXE%
goto erro
:erro_final_dist
echo Falha ao criar/copiar o executavel para a pasta final.
goto erro
:erro_exe_final
echo O executavel nao esta presente na pasta final.
goto erro
:erro

echo.
echo ============================================================
echo FALHA DURANTE A GERACAO DO EXECUTAVEL.
echo ============================================================
echo.
echo Area de staging: %STAGE%
echo Envie o log completo desta janela.
echo.
pause
exit /b 1
