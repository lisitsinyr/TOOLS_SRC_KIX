'=======================================================================================
' APPLinks.vbs
'=======================================================================================

   Option Explicit

   '=======================================================================================
   const conFileNameINI = "links.ini"
   const conFileNameINIAPP = "apps.ini"
   const conShablon     = "CommandFile|Arg|WorkDir|IconFile|DestDir|SubDir|"
   '=======================================================================================

   const conFPrograms = "Программы_TEST"
   const conSubDir = ""
   const conFTasks = "Задачи_TEST"
   const conSubDirTasks = ""

   Dim   strFileIniAPPs

   Dim   strPrograms
   Dim   strTasks

   '=======================================================================================
   Dim   strFPrograms, strFTasks
   Dim   strSPrograms, strSTasks
   Dim   strSubDir, strSubDirTasks
   '=======================================================================================
   Dim   strTempDir 
   Dim   strWorkDir 
   Dim   strCmdFile 
   Dim   strNameLnkT
   Dim   strNameLnkP
   Dim   strIconFile
   Dim   strDestSubDir 
   Dim   strDestDirWork
   Dim   strDestDirWork1
   Dim   strDestDirWork2
   '=======================================================================================
   Dim   strLogFile
   Dim   strLogDir

'--------------------------------------------------------------------
' Constructor
'--------------------------------------------------------------------
'begin
   strFileIniAPPs = sstrVBLDir & "\" & conFileNameINIAPP
   strPrograms = conFPrograms & conSubDir
   strTasks = conFTasks & conSubDirTasks
'end constructor

'------------------------------------------------------
' TitleInfo
'------------------------------------------------------
function TitleInfo (argFile)
   Dim FileName, objFile, strGroup
'begin
   PrintGeneralTitle argFile

   LogAdd argFile, "I", "TServer       = " + sstrLoginTerminalServer
   LogAdd argFile, "I", "APPLinks      = " + strFileIniAPPs

   ListDrives(argFile)

   LogAdd argFile, "I", "==========================================="
   LogAdd argFile, "I", " Группы пользователя"
   LogAdd argFile, "I", "==========================================="
   SetEcho(False)
   For Each strGroup In sarrUserAllGroups
      LogAdd argFile, "I", strGroup
   next
   SetEcho(True)
End Function

'-------------------------------------------------
' LogonLog
'-------------------------------------------------
function LogonLog (argDir)
   Dim strFile, objFile, strDir
'begin
   '------------------------------------------------------
   if Ucase(sstrWKSTA) <> "APPS2" then
      strFile = argDir+"\"+AddCharR(" ",sstrWKSTA,12)+" "+"("+Trim(Kix_CPU)+")"+"."+AddCharR(" ",Kix_ProductType,25)
      sobjFSO.DeleteFile (strFile)
      Set objFile = sobjFSO.CreateTextFile(strFile, True)
      objFile.WriteLine("")
      objFile.Close
   end if
   '------------------------------------------------------
   strDir = argDir+"\logon"
   if not sobjFSO.FolderExist(strDir) then
      MD (strDir)
   end if
   strFile = LogFileName (2, strDir+"\logon", sstrUserID & ".log")
   LogAdd strFile, "P", "|" & sstrUserID & "|" & Kix_FullName & "|" &       _
                              Kix_WKSTA & "|" & Kix_ProductType & "|" &     _
                              sstrUserDomain & "|" & Kix_LServer & "|" &    _
                              sstrWKSTA & "|" & sstrPCDomain
   '------------------------------------------------------
End Function

'------------------------------------------------------
' GetDirWork
'------------------------------------------------------
function GetDirWork (argDir, argSubDirAPP)
'begin
WScript.Echo argDir & "/" & argSubDirAPP
   if UCase(argDir) = UCase(strFPrograms) then
      GetDirWork  = sstrDesktop & "\" & strPrograms
   elseif UCase(argDir) = UCase(strFTasks) then
      GetDirWork  = sstrDesktop & "\" & strTasks
   elseif UCase(argDir) = "DESKTOP" then
      GetDirWork  = sstrDesktop
   elseif UCase(argDir) = "PROGRAMS" then
      GetDirWork  = sstrPrograms
   elseif UCase(argDir) = "STARTMENU" then
      GetDirWork  = sstrStartMenu
   else
      if sobjFSO.FolderExists(argDir) then
         GetDirWork = argDir
      else
         GetDirWork  = sstrDesktop & "\" & strPrograms
      end if
   End if

   if argSubDirAPP <> "" then
      GetDirWork  = GetDirWork & "\" & argSubDirAPP
   end if

End Function

'------------------------------------------------------
' ClearDir
'------------------------------------------------------
function ClearDir (argDir, argLogFile)
   Dim strS
'begin
   strS = GetDirWork (argDir, "")
   LogAdd argLogFile, "I", "Удаление ... " & strS

   'DelDir (strS, argLogFile)
   'Rd (strS)

   '----------------------------------------------------------
   ' Start Menu
   '----------------------------------------------------------
   if UCase(argDir) = UCase(strFPrograms) then
      strS = sstrPrograms & "\" & strPrograms
      LogAdd argLogFile, "I", "Удаление ... " & strS

      'DelDir(strS, argLogFile)

   elseif UCase(argDir) = UCase(strFTasks) then
      strS = sstrPrograms & "\" & strTasks
      LogAdd argLogFile, "I", "Удаление ... " & strS

      'DelDir (strS, argLogFile)

   End if
End Function





'------------------------------------------------------
' CreateLinks3 ($FileIni)
'------------------------------------------------------
function CreateLinks3 (argFileIni)
   Dim arrDomains, arrAPPs, strAPP, strDomainAPP
'begin 
   arrDomains = ReadSections(argFileINI)
   for each strDomain in arrDomains
      strDomainAPP = strDomain
      if strDomain <> ""
         LogAdd strLogFile, "I", "====================================="
         LogAdd strLogFile, "I", "Домен: " & strDomainAPP
         LogAdd strLogFile, "I", "====================================="
         arrAPPs = ReadSection (argFileINI, strDomain)
         for each strAPP in strAPPs
            if strAPP <> ""
               ' LogAdd strLogFile, "I", "-------------------------------------"
               LogAdd argLogFile,"I","Программа: " & strAPP
               ' LogAdd strLogFile, "I", "-------------------------------------"
               ' CreateLinksAPP (argFileIni, $Domain, $APP)
            end if
         next
      end if
   next
End Function

'------------------------------------------------------
' CreateLinks
'------------------------------------------------------
function CreateLinks
   Dim strS 
'begin
   On Error Resume Next
   '------------------------------------------------------
   LogonLog(sconLogPath)
   '------------------------------------------------------
   if Kix_LDomain = "" then
      strLogDir = sconLogPath+"\logon"
   else
      strLogDir = sconLogPath+"\logon\"+Kix_LDomain
   end if
   if not sobjFSO.FolderExists(strLogDir) then 
      'sobjFSO.CreateFolder (strLogDir) 
      MD (strLogDir)
   end if

   if Err.Number <> 0 then
      WScript.Echo "CreateLinks: " & Err.Number & " " & Err.Description & " " & " "
      Err.Clear
   end if

   strS = AddCharR("_", sstrUserID, 8)
   strLogFile = strS & "_" & UCase(Kix_WKSTA) & ".lnk"
   strLogFile = LogFileName (1, strLogDir, strLogFile)

   TitleInfo (strLogFile)
   '------------------------------------------------------
   On Error GoTo 0 

   LogAdd strLogFile, "I", "==========================================="
   LogAdd strLogFile, "I", " Создание ярлыков "
   LogAdd strLogFile, "I", "==========================================="

   ClearDir strPrograms, strLogFile
   ClearDir strTasks, strLogFile

   strTempDir  = sobjWshSysEnv_USER("TEMP")
   strWorkDir  = "\\fsgu\app"
   strCmdFile  = strWorkDir & "\Tools\CreateLinks.bat"
   strNameLnkT = "Обновить список задач"
   strNameLnkP = "Обновить список программ"
   strIconFile = strWorkDir & "\Tools\ico\Lightning.ico"

   WScript.Echo strTempDir & "/" & strWorkDir & "/" & strCmdFile & "/" & strNameLnkP & "/" & strIconFile

   if (IsTerminalApplication) then
      strDestSubDir  =  strTasks
      strDestDirWork = GetDirWork(strTasks, "")

      WScript.Echo strCmdFile & "/" & strTempDir & "/" & strIconFile & "/" & strNameLnkT & "/" & strDestDirWork & "/" & strDestSubDir

      'CreateOne (strCmdFile, "", strTempDir, strIconFile, strNameLnkT, strDestDirWork, strDestSubDir)

   else
      '------------------------------------------------------
      strDestSubDir  = strPrograms
      strDestDirWork = GetDirWork(strPrograms, "")

      WScript.Echo strCmdFile & "/" & strTempDir & "/" & strIconFile & "/" & strNameLnkT & "/" & strDestDirWork & "/" & strDestSubDir

      'CreateOne (strCmdFile, "", strTempDir, strIconFile, strNameLnkP, strDestDirWork, strDestSubDir)

      '------------------------------------------------------
      strDestSubDir  = "Programs" & "\" & strPrograms
      strDestDirWork = sstrPrograms & "\" & strPrograms

      WScript.Echo strCmdFile & "/" & strTempDir & "/" & strIconFile & "/" & strNameLnkT & "/" & strDestDirWork & "/" & strDestSubDir

      'CreateOne (strCmdFile, "", strTempDir, strIconFile, strNameLnkP, strDestDirWork, strDestSubDir)

      '------------------------------------------------------
      'strCmdFile  = WorkDir+"\Tools\Remote Desktop\mstsc.exe"
      'strIconFile = WorkDir+"\Tools\Remote Desktop\mstsc.exe"
      'if sstrDomainUser = "DSP"
      '   strNameLnkP = "Remote_DSP"
      '   strParamRDP = '"'+strWorkDir+'\Tools\Remote Desktop\apps2_DSP.rdp"'
      'else
      '   strNameLnkP = "Remote_GU"
      '   strParamRDP = '"'+strWorkDir+'\Tools\Remote Desktop\apps2.rdp"'
      'end if
      'strDestDirWork = strDesktop
      'strDestSubDir  = strDesktop
      'strDestDirWork = Programs
      'strDestSubDir  = strPrograms
      '------------------------------------------------------
      ' CreateOne (strCmdFile, strParamRDP, strTempDir, strIconFile, strNameLnkP, strDestDirWork, strDestSubDir)
      '------------------------------------------------------
   end if

   '------------------------------------------------------
   ' 
   '------------------------------------------------------
   'CreateLinks3 (sstrFileIniAPPs)

   '------------------------------------------------------
   ' CopyFolder
   '------------------------------------------------------
   if (not IsTerminalApplication) then
      strDestDirWork1 = GetDirWork(strPrograms, "")
      strDestDirWork2 = strPrograms & "\" & strPrograms
      LogAdd strLogFile, "I", strDestDirWork1 & " => " & strDestDirWork2
      'sobjFSO.CopyFolder (strDestDirWork1, strDestDirWork2, True)
      'BacFiles (strDestDirWork1, strDestDirWork2, "*.*", True)
   end if

End Function

