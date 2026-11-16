@echo off
setlocal DisableDelayedExpansion
title Remover painel CMYK e Pantone
rem Mantem o console aberto mesmo se o PowerShell falhar antes de executar.
if /i "%~1"=="--console" goto :executar
"%SystemRoot%\System32\cmd.exe" /d /s /k ""%~f0" --console"
exit /b

:executar
echo Remover painel CMYK e Pantone
echo.
echo Esta janela permanece aberta. Feche pelo X quando terminar.
echo Extraia TODO o ZIP antes de executar este arquivo.
echo.
if not exist "%~dp0Instalar.ps1" goto :incompleto
if not exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" goto :sem_powershell
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Instalar.ps1" -Remover
set "resultado=%errorlevel%"
echo.
echo Finalizado. Codigo de retorno: %resultado%
echo Leia a mensagem acima. Se houve erro, envie uma foto desta janela.
echo.
exit /b %resultado%

:incompleto
echo Nao encontrei Instalar.ps1 ao lado deste arquivo.
echo Feche, clique com o botao direito no ZIP e escolha Extrair Tudo.
echo Depois abra a pasta extraida e execute Instalar.cmd como administrador.
echo.
exit /b 2

:sem_powershell
echo Nao foi encontrado o Windows PowerShell no caminho esperado.
echo Envie uma foto desta janela para verificarmos.
echo.
exit /b 3
