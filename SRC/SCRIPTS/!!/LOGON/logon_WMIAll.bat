@Echo off
echo "=================================="
echo "COMPUTERNAME  " %COMPUTERNAME%
echo "USERNAME      " %USERNAME%
echo "USERDOMAIN    " %USERDOMAIN%
echo "USERDNSDOMAIN " %USERDNSDOMAIN%
echo "OS            " %OS%
echo "=================================="

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
set KXLDIR=D:\tools
goto Stop
:Stop
echo Значение переменной среды KXLDIR=%KXLDIR%
rem -------------------------------------------------------------------

:StepBegin
rem \\S73FS01\APPInfo\tools\kix\KIX_4_53\kix32.exe -f \\S73FS01\APPInfo\tools\Logon.kix
kix32.exe -f %KXLDIR%\logon_WMIALL.kix
:StepEnd
