'=======================================================================================
' Logon.kix
'=======================================================================================

'=======================================================================================
' Uses
'=======================================================================================

   Dim strLogDir, strLogFile, s
   Dim strKey

   Dim strFileINI

   Dim sstrVBLDir

   Dim objWshSysEnv_USER

   Dim sobjWshShell

   Set sobjWshShell = WScript.CreateObject ("WScript.Shell")

   Set objWshSysEnv_USER = sobjWshShell.Environment("USER")
   if objWshSysEnv_USER("VBSDIR") <> "" then
      sstrVBLDir = objWshSysEnv_USER("VBSDIR")
   else
      sstrVBLDir = sconVBLDir
   end if
   WScript.Echo sstrVBLDir

   LoadLib1 (sstrVBLDir & "\VBL\LUConst.vbs")
   LoadLib1 (sstrVBLDir & "\APPLinks.vbs")


'begin
   SetEcho (True)   

   On Error Resume Next
   '------------------------------------------------------
   if Kix_LDomain = "" then
      strLogDir = sconLogPath+"\logon"
   else
      strLogDir = sconLogPath+"\logon\"+sstrUserDomain
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
   strLogFile = strS & "_" & UCase(Kix_WKSTA) & ".log"
   strLogFile = LogFileName (1, strLogDir, strLogFile)
   '------------------------------------------------------
   On Error GoTo 0 

   '-----------------------------------------------------------
   ' Синхронизация времени
   '-----------------------------------------------------------
   'LogAdd (LogFile, "I", "===========================================")
   'LogAdd (LogFile, "I", "Синхронизация времени с " + Kix_LServer+"...")
   'SetTime "Kix_LServer"
   '-----------------------------------------------------------
   ' Установка начальных настроек для Internet Explorer
   '-----------------------------------------------------------
   'LogAdd (LogFile, "I", "===========================================")
   'LogAdd (LogFile, "I", "Установка начальных настроек для Internet Explorer"+"...")
   'strKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
   'WriteValue ($Key, "ProxyOverride", "www.ulnsk.cbr.ru'ulnsk.cbr.ru'<local>", "REG_SZ")
   'WriteValue ($Key, "ProxyServer",   "proxy.ulnsk.cbr.ru:3128" ,              "REG_SZ")
   'WriteValue ($Key, "ProxyEnable",   "0,0,0,1",                               "REG_BINARY")
   'strKey = "HKCU\Software\Microsoft\Internet Explorer\Main"
   'WriteValue ($Key, "Start Page", "http://www.ulnsk.cbr.ru",                  "REG_SZ")
   '-----------------------------------------------------------
   ' Подключение сетевых дисков
   '-----------------------------------------------------------
   WScript.Echo strFileINI
   WScript.Echo strLogFile

   MapDrives (strLogFile)

   'ListDrives (strLogFile)
   '-----------------------------------------------------------
   ' 
   '-----------------------------------------------------------
   if UCase(Kix_WKSTA) = UCase("S73ap01") or            _
      UCase(Kix_WKSTA) = UCase("S73ap02") or            _
      UCase(Kix_WKSTA) = UCase("S73ap03") or            _
      UCase(Kix_WKSTA) = UCase("S73ap04") or            _
      UCase(Kix_WKSTA) = UCase("S73ap05") then
         '-----------------------------------------------------------
         ' Выполнение скрипта для S73ap01,S73ap02,S73ap03,S73ap04,S73ap05
         '-----------------------------------------------------------
         LogAdd strLogFile, "I", "Выполнение скрипта для терминала: " & UCase(Kix_WKSTA)

         'if not Exist("w:\TEMP") then MD w:\TEMP end if 

         'Shell "regedit -s G:\Tools\FontCitrix.reg"
   else
      '-----------------------------------------------------------
      ' Создание ярлыков для локальных машин
      '-----------------------------------------------------------
      'CreateLinks
   end if
'end


                                                                                                                                                        
'-----------------------------------------------------------
' MapDrives
'-----------------------------------------------------------
function MapDrives (argLogFile)
   Dim blnDone, i, j, n, s, arrObjects, strObject, strItem, strLDomen, arrUserAllGroups
   Dim arrDrives, strDrive, strPath, intPersistent, arrIndex
'begin
   LogAdd argLogFile, "I", "==========================================="
   LogAdd argLogFile, "I", "Подключение сетевых дисков для " & Kix_UserID & "..."

   WScript.Echo argLogFile

   '------------------------------------------------------------------------
   arrUserAllGroups = GetDomainUserGlobalGroups (False, sstrUserDomain, sstrUserID)
   '------------------------------------------------------------------------

   arrObjects = ReadSections_LU (strFileINI)

   for each strObject in arrObjects
      j = 0
      if strObject <> "" then
         LogAdd argLogFile, "I", "[" & strObject & "]"
         n = WordCount(strObject, ",")
         if n > 0 then
            blnDone = False
            i = 1
            do while (i <= n) and (blnDone = 0)
               strItem = ExtractWord(i, strObject, ",")
               if strItem <> "" then
                  '------------------------------------------
                  ' User без домена
                  '------------------------------------------
                  if WordCount(strItem, "\") = 1 then
                     strItem = sstrDOMEN_Default & "\" & strItem & "\"
                  end if
                  strLDomen = ExtractWord(1, strItem, "\")
                  s = ExtractWord(2, strItem, "\")

                  if (UCase(s) = UCase(Kix_UserID)) or (UCase(s) = UCase("All")) then
                     blnDone = True
                  else
                     if WordCount(strItem, "\") = 3 then
                        '-----------------------------------------------------------
                        ' LogAdd argLogFile, "I", "Это пользователь = " & strItem
                        '-----------------------------------------------------------
                     else
                        '-----------------------------------------------------------
                        ' LogAdd argLogFile, "I", "Это группа = " & strItem)
                        '-----------------------------------------------------------
                        'if (UCase(sstrUserDomain) = "GU"    and UCase(strLDomen) = "DSP") or _
                        '   (UCase(sstrUserDomain) = "DSP"   and UCase(strLDomen) = "GU")  or _
                        '   (UCase(sstrUserDomain) = "ULNSK" and UCase(strLDomen) = "DSP") or _
                        '   (UCase(sstrUserDomain) = "ULNSK" and UCase(strLDomen) = "GU")  then
                        '   '-----------------------------------------------------------
                        '   ' LogAdd argLogFile, "I", "Нет проверки вхождения в группу!" & strItem & ". (" & Kix_LDomain & "=" & strLDomen & ")"
                        '   '-----------------------------------------------------------
                        'else
                           'If InGroup (strItem) then
                           '   blnDone = True
                           'else
                              'if IndexOf (arrUserAllGroups, strItem) >= 0 then

                              arrIndex = Filter(arrUserAllGroups, strItem)

                              if UBound(arrIndex) >= 0 then
                                 blnDone = True
                              end if
                           'end if
                        'end if
                     end if
                  end if

               end if
               i = i + 1
            loop
            '------------------------------------------------------
            if blnDone = True then

               arrDrives = ReadSection_LU(strFileINI, strObject)

               for each strDrive in arrDrives
                  s = ReadString (strFileIni, strObject, strDrive, "") & "|"
                  strPath = ExtractWord(1, s, "|")
                  intPersistent = ExtractWord(2, s, "|")
                  LogAdd argLogFile, "I", strDrive & " => " & strPath
                  MapDrive Left(strDrive,1), strPath, argLogFile, intPersistent
               next
            end if
            '------------------------------------------------------
         end if
      end if
   next
end function

'=======================================================================================
' Подключение внешних библиотек
' strFullPatchLib - подключаемая библиотека (с полным путем)
' strCode - данные для подключения 
' после выхода из функции необходимо выполнить:
' Execute strCode
'=======================================================================================
Function LoadLib1 (argFullPatchLib)
   Dim objFSO, objTextStream
'begin
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   LoadLib1 = ""
   On Error Resume Next
   Set objTextStream = objFSO.OpenTextFile(argFullPatchLib, 1)
   if Err.Number <> 0 then
      WScript.Echo "LoadLib: " & Err.Number & " " & Err.Description & " " & argFullPatchLib 
      Err.Clear
   else
      LoadLib1 = objTextStream.ReadAll
      objTextStream.Close
      ExecuteGlobal LoadLib1
   end if
   On Error GoTo 0 
End Function

