@echo off
rem -------------------------------------------------------------------
if "%KXLDIR%" == "" goto Begin
goto Stop
:Begin
echo Значение переменная среды KXLDIR не установлено
if "%COMPUTERNAME%" == "%USERDOMAIN%" goto Local
:Network
set KXLDIR=\\S73FS01\APPInfo\tools
goto Stop
:Local
set KXLDIR=c:\tools
goto Stop
:Stop
echo Значение переменная среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\bac.kix "$Lib=%1" "$Func=%2" "$P1=%3" "$KxlDir=%KXLDIR%" "$Debug=0"

if "%2" == "BAT" goto syncTerm
goto end

:syncTerm

echo empty > \\S73FS01\APPInfo\Tools\syncTerm\sync.
rem echo empty > \\FSGU\APP\Tools\syncTerm\sync.

:end
