@echo on
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
set KXLDIR=D:\TOOLS
goto Stop
:Stop
echo Значение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

kix32.exe %KXLDIR%\ListDir.kix "$KxlDir=%KXLDIR%" "$Format=%1" "$NLevel=%2"
