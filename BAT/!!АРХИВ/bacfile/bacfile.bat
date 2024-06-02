@echo off
rem -------------------------------------------------------------------
if "%KXLDIR%" == "" goto Begin
goto Stop
:Begin
echo Значение переменной среды KXLDIR не установлено
if "%COMPUTERNAME%" == "%USERDOMAIN%" goto Local
:Network
set KXLDIR=\\S73FS01\APPInfo\tools
goto Stop
:Local
set KXLDIR=c:\tools
goto Stop
:Stop
echo Значение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\bacfile.kix "$S=%1" "$D=%2" "$M=%3" "$Log=%4" "$KxlDir=%KXLDIR%"
