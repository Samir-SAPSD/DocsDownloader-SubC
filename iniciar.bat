@echo off
chcp 65001 > nul
echo ==========================================
echo Executando a aplicação...
echo ==========================================
echo Tentando executar com Python...
call python src/downloadFiles.py
if %errorlevel% equ 0 goto :success

echo.
echo Execução com Python falhou. Tentando via Anaconda (caminho absoluto)...
call "C:\ProgramData\anaconda3\python.exe" src/downloadFiles.py
if %errorlevel% neq 0 (
    echo Erro na execução do script Python (Python e Anaconda falharam).
    goto :error
)

:success
echo.
echo Aplicação finalizada com sucesso.
pause
exit /b 0

:error
echo.
echo ==========================================
echo OCORREU UM ERRO DURANTE A EXECUÇÃO
echo ==========================================
echo A janela será fechada após pressionar uma tecla...
pause
exit /b 1