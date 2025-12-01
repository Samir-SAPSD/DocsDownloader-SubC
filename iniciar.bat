@echo off
chcp 65001 > nul

:: 1. Verifica se o Python no PATH tem as dependências
python -c "import customtkinter" >nul 2>nul
if %errorlevel% equ 0 (
    start "" pythonw src/downloadFiles.py
    exit
)

:: 2. Se falhou, tenta o Anaconda (caminho absoluto)
if exist "C:\ProgramData\anaconda3\pythonw.exe" (
    start "" "C:\ProgramData\anaconda3\pythonw.exe" src/downloadFiles.py
    exit
)

:: 3. Se chegou aqui, erro
echo.
echo ==========================================
echo ERRO: Não foi possível iniciar a aplicação.
echo ==========================================
echo Verifique se o Python está instalado e se 'customtkinter' está presente.
pause
exit /b 1