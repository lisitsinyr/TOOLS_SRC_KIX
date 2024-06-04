@Echo off
echo "=================================="
echo "COMPUTERNAME  " %COMPUTERNAME%
echo "USERNAME      " %USERNAME%
echo "USERDOMAIN    " %USERDOMAIN%
echo "USERDNSDOMAIN " %USERDNSDOMAIN%
echo "OS            " %OS%
echo "=================================="

:StepBegin

\\S73FS01\APPInfo\tools\kix\KIX_4_53\kix32.exe -f \\S73FS01\APPInfo\tools\Logon.kix

:StepEnd
