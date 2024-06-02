rem arj a -r -a %1 %1
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

%KXLDIR%\params %1 > %TEMP%\ParamStr.kxl
kix32.exe %KXLDIR%\PackDir.kix $FileParamStr=%TEMP%\ParamStr.kxl $ARC=arj $KxlDir="%KXLDIR%"

