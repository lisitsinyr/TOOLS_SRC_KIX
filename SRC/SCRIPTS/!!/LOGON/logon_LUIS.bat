@Echo off
echo "=================================="
echo "COMPUTERNAME  " %COMPUTERNAME%
echo "USERNAME      " %USERNAME%
echo "USERDOMAIN    " %USERDOMAIN%
echo "USERDNSDOMAIN " %USERDNSDOMAIN%
echo "OS            " %OS%
echo "=================================="

if "%OS%"=="Windows_NT" goto WindowsNT
if "%OS%"==""           goto Windows9x
goto StepEnd

rem ---------------------------------------------
rem WindowsNT
rem ---------------------------------------------
:WindowsNT
rem echo "net use G: /delete"
rem net use G: /delete
rem echo "net use G: \\FSGU\APP /persistent:yes"
rem net use G: \\FSGU\APP /persistent:yes
rem echo "net use W: /delete"
rem net use W: /delete
goto StepBegin

rem ---------------------------------------------
rem Windows9x
rem ---------------------------------------------
:Windows9x
echo "net use G: /delete /yes"
net use G: /delete /yes
echo "net use G: \\FSGU\APP /yes"
net use G: \\FSGU\APP /yes
goto StepBegin

:StepBegin

echo ===============================================
echo \\S73FS01\APPInfo\tools\Logon.kix
echo ===============================================
\\S73FS01\APPInfo\tools\kix\KIX_4_53\kix32.exe -f \\S73FS01\APPInfo\tools\Logon.kix

:StepEnd
