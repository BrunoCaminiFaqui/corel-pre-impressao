@echo off
setlocal DisableDelayedExpansion
title Registro da instalacao do painel
if /i "%~1"=="--console" goto :abrir
"%SystemRoot%\System32\cmd.exe" /d /s /k ""%~f0" --console"
exit /b
:abrir
if not exist "%LOCALAPPDATA%\Imprell\SeparacoesCorel\instalacao-*.log" goto :nenhum
"%SystemRoot%\explorer.exe" "%LOCALAPPDATA%\Imprell\SeparacoesCorel"
echo Aberta a pasta dos registros. Envie o arquivo instalacao-...log mais recente.
exit /b 0
:nenhum
echo Nao ha registro de instalacao nesta conta do Windows.
echo Extraia o ZIP e execute Instalar.cmd; envie uma foto da janela se houver erro.
exit /b 1
