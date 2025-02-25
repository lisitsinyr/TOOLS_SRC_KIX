
'------------------------------------------------------
' CreateINI($FileINI)
'------------------------------------------------------
function CreateINI($FileINI)
   Dim arrDomains
'begin
   arrDomains = GetDomains
   for each Domain in Domains
      if (UCase(Domain) <> "WORKGROUP") and (UCase($Domain) <> "UI") and  (UCase(Domain) <> "UINF")
         Groups = GetLocalGroups($Domain)
         for $i=0 to UBound($Groups,2)
            $Path = $Groups[1,$i]
            $Path = ""
            $s0 = $Domain+"\"+$Groups[0,$i]+"="+$Path
            if not ValueExists ($FileINI, $Domain, $Groups[0,$i])
               ' нет в links.ini
               $s = "New="+$s0
               LogAdd($Log,$LogFile,"I", $s)
               if ($Path <> "") and Exist($Path)
                  $=WriteString ($FileINI, $Domain, $Groups[0,$i], $Path)
               else
                  $=WriteString ($FileINI, $Domain, $Groups[0,$i], "\\fsgu\app\")
               end if
            else
               ' есть в links.ini
               if ($Path <> "") and Exist($Path)
                  $=WriteString ($FileINI, $Domain, $Groups[0,$i], $Path)
               else
                  $Path = ReadString ($FileINI, $Domain, $Groups[0,$i], "")
                  $s0 = $Domain+"\"+$Groups[0,$i]+"="+$Path
               end if
               $s = "Old="+$s0
               LogAdd($Log,$LogFile,"I", $s)
            end if
         next
      end if
   next
End Function

'-------------------------------------------------------------------------------
' CreateLinksIni
'-------------------------------------------------------------------------------
function CreateLinksIni($Mask,$SubDirAPP,$FileINIApp,$Group)
   Dim $LFile,$LPath,$LFullFileName
'begin
   $LPath = GetFilePath($Mask)
   if $LPath <> ""
      $LPath = GetFilePath($Mask)+"\"
   else
      $LPath = @CurDir+"\"
   end if
   $LFile = Dir ($Mask)
   WHILE @ERROR = 0 AND $LFile
      IF $LFile <> "." AND $LFile <> ".."
         $LFullFileName = $LPath+$LFile
         IF not (GetFileAttr ($LPath+$LFile) & 16)    ' is it a directory ?
            $Key = Substr($LFile,1,Instr($LFile,".")-1)
            $Value = $LFullFileName "|" GetFilePath($LFullFileName) "||" $LFullFileName "|Программы|$SubDirAPP|"
            $Result = WriteString($FileINIApp, $Group, $Key, $Value)
         end if
      end if
      if @ERROR = 0
         $LFile = Dir("")
      end if
   loop
End Function

'------------------------------------------------------
' CreateINIAPP($FileINI)
'------------------------------------------------------
function CreateINIAPP($FileINI)
   Dim $FileINIApp, $Domains, $Groups
'begin
   $Domains = ReadSections($FileINI)
   for each $Domain in $Domains
      $Groups = ReadSection($FileINI, $Section)
      for each $Group in $Groups
         $FileINIApp = ReadString ($FileINI, $Domain, $Group, "")+"\"+strFileNameINI
         if Exist($FileINIApp) 
            if SectionExists($FileINIApp, $Group) = False
               ' $Result = WriteString($FileINIApp, $Group, "name", $Shablon)
               CreateLinksIni("*.bat", "", $FileINIApp, $Group)
            end if
         else
            ' $Result = WriteString($FileINIApp, $Group, "name", $Shablon)
            CreateLinksIni("*.bat", "", $FileINIApp, $Group)
         end if
      next
   next
End Function

'------------------------------------------------------
' CreateINIAPP_OLD($FileINI)
'------------------------------------------------------
function CreateINIAPP_OLD ($FileINI)
   Dim $FileINIApp, $Domains, $Groups
'begin
   $Domains = ReadSections($FileINI)
   for each $Domain in $Domains
      $Groups = ReadSection($FileINI, $Domain)

      for each $Group in $Groups
         $FileINIApp = ReadString ($FileINI, $Domain, $Group, "")+"\"+strFileNameINI
         if Exist($FileINIApp) 
            if SectionExists($FileINIApp, $Group) = False
               $Result = WriteString($FileINIApp, $Group, "name", $Shablon)
            end if
         else
            $Result = WriteString($FileINIApp, $Group, "name", $Shablon)
         end if
      next
   next
End Function




















'------------------------------------------------------
' CreateOne ($CommandFile, $Arg, $WorkDir, $IconFile, $Name, $DestDirWork, $DestSubDir, optional $WindowStyle, optional $PathAPP)
'------------------------------------------------------
function CreateOne ($CommandFile, $Arg, $WorkDir, $IconFile, $Name, $DestDirWork, $DestSubDir, optional $WindowStyle, optional $PathAPP)
   Dim $s,$FileNamePif,$FileNameLnk
'begin
   $FileNamePif = $DestDirWork+"\"+$Name+".pif"
   $FileNameLnk = $DestDirWork+"\"+$Name+".lnk"
   if blnDebug 
      $s = $DestDirWork+"\"+$Name+"|"+$CommandFile+"|"+$Arg+"|"+$WorkDir+"|"+$IconFile+"|"+$Name+"|"+$WindowStyle
      LogAdd($Log,$LogFile,"I", $s)
   end if
   $FileExt     = GetFileExt($CommandFile)

   '----------------------------------------------------
   ' 1 вариант
   '----------------------------------------------------
   '$Minimize = 0
   '$Replace = 0
   '$RunInOwnSpace = 0
   '$Result = CreateLink ($CommandFile, $Name, $IconFile, $IconIndex, $WorkDir, $Minimize, $Replace, $RunInOwnSpace)
   '----------------------------------------------------

   '----------------------------------------------------
   ' 2 вариант
   '----------------------------------------------------
   $Result = wshShortCut ($DestDirWork+"\"+$Name, $CommandFile, $Arg, $WorkDir, $IconFile, $WindowStyle, $Name)
   if @Error <> 0
      LogAdd($Log,$LogFile,"I", "Ошибка при создании ("+$Name+"): "+@Error+" "+$SError) 
      '----------------------------------------------------
      ' 1 вариант
      '----------------------------------------------------
      '$Minimize = 0
      '$Replace = 0
      '$RunInOwnSpace = 0
      '$Result = CreateLink ($CommandFile, $Name, $IconFile, $IconIndex, $WorkDir, $Minimize, $Replace, $RunInOwnSpace)
      'CreateLinkLU ($KxlDir+"\Links.exe", $DestDirWork, $name, $CommandFile, $arg, $WorkDir, $IconFile, , )
      '----------------------------------------------------
   else
      $ExeDir = $KxlDir+"\"

      '----------------------------------------------------------------------
      ' PIF
      '----------------------------------------------------------------------
      if Exist($FileNamePif)
         if Exist($PathAPP+"\"+$Name+".pif")
            LogAdd($Log,$LogFile,"I", $PathAPP+"\"+$Name+".pif"+" => "+$DestDirWork) 
            Copy $PathAPP+"\"+$Name+".pif" $DestDirWork
         else
            select
               case (@ProductType = "Windows 98") or (@ProductType = "Windows 95")
                  '--------------------------------------------------------------
                  ' +0063h 0000h - default
                  '        0010h - Закрывать окно по завершении работы
                  '--------------------------------------------------------------
                  $s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNamePif+'"'+" -o $63 -d $0010"
                  if blnDebug 
                     LogAdd($Log,$LogFile,"I", $s) 
                  end if
                  shell "$s"
            End Select

            '--------------------------------------------------------------
            ' 01ADh 00 02 10 02h - default
            '                08h - полноэкранный режим
            '          00        - обчный размер окна
            '          10        - свернутое в значек
            '          20        - развернутое во весь экран
            '          80        - режим MS DOS        
            '--------------------------------------------------------------
            if $WindowStyle = 3
               $s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNamePif+'"'+" -o $1ad -d $00221008"
               if blnDebug LogAdd($Log,$LogFile,"I", $s) end if
               shell "$s"
            end if
         end if

      end if

      '----------------------------------------------------------------------
      ' LNK
      '----------------------------------------------------------------------
      if Exist($FileNameLnk)
         if Exist($PathAPP+"\"+$Name+".lnk")
            LogAdd($Log,$LogFile,"I", $PathAPP+"\"+$Name+".lnk"+" => "+$DestDirWork) 
            Copy $PathAPP+"\"+$Name+".lnk" $DestDirWork
         else
            if $WindowStyle = 3
               '--------------------------------------------------------------
               ' Отображение во весь экран
               '--------------------------------------------------------------
               if UCase($FileExt) = "EXE"
                  $s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $3c -d $03"
                  if blnDebug 
                     LogAdd($Log,$LogFile,"I", $s) 
                  end if
                  ' shell "$s"
               end if
               '--------------------------------------------------------------
               ' Отображение во весь экран
               '--------------------------------------------------------------
               if UCase($FileExt) = "BAT"
                  select
                     case $W2000
                        $s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $234 -d $01"
                        shell "$s"
                     case $WXP or $WNT
                        $s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $12d -d $1c"
                        shell "$s"
                  End Select
                  '--------------------------------------------------------------
                  '$s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $131 -d $2d"
                  'if blnDebug LogAdd($Log,$LogFile,"I", $s) end if
                  'shell "$s"
                  '--------------------------------------------------------------
               end if

               '--------------------------------------------------------------
               ' Размер буфера экрана (высота)
               '--------------------------------------------------------------
               '$s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $23f -d $19"
               'if blnDebug LogAdd($Log,$LogFile,"I", $s) end if
               'shell "$s"
               '--------------------------------------------------------------
               ' Размер окна (высота)
               '--------------------------------------------------------------
               '$s = $ExeDir+"ChangeFile.exe -f "+'"'+$FileNameLnk+'"'+" -o $243 -d $19"
               'if blnDebug LogAdd($Log,$LogFile,"I", $s) end if
               'shell "$s"
               '--------------------------------------------------------------
            end if
         end if
      end if
   end if

   '----------------------------------------------------
   ' 3 вариант
   '$Result = CreateLinkLU ($KxlDir+"\Links.exe", $DestDirWork, $Name, $CommandFile, $Arg, $WorkDir, $IconFile,, $Name)
   '----------------------------------------------------

   '----------------------------------------------------
   $s = "      Создан link "
   $s = "      "
   if Exist($FileNamePif)
      $s = $s + $DestSubDir+"\"+GetFileName($FileNamePif)
   else
      $s = $s + $DestSubDir+"\"+GetFileName($FileNameLnk)
   end if
   if blnDebug 
      LogAdd($Log,$LogFile,"I",$s+". Ok!")
   else
      LogAdd(10,$LogFile,"I",$s+". Ok!")
   end if
   '----------------------------------------------------
End Function

'------------------------------------------------------
' DeleteOne ($Name, $DestDirWork, $DestSubDir)
'------------------------------------------------------
function DeleteOne ($Name, $DestDirWork, $DestSubDir)
   Dim $s,$FileNamePif,$FileNameLnk
'begin
   $FileNamePif = $DestDirWork+"\"+$Name+".pif"
   $FileNameLnk = $DestDirWork+"\"+$Name+".lnk"
   $s = "Delete link ... "
   if Exist($FileNamePif)
      $s = $s + $DestSubDir+"\"+GetFileName($FileNamePif)
      Del "$FileNamePif"
      LogAdd($Log,$LogFile,"I",$s)
   end if
   if Exist($FileNameLnk)
      $s = $s + $DestSubDir+"\"+GetFileName($FileNameLnk)
      Del "$FileNameLnk"
      LogAdd($Log,$LogFile,"I",$s)
   end if
End Function

'------------------------------------------------------
' CreateLinksGroup ($FileIni, $Group, $APP)
'------------------------------------------------------
function CreateLinksGroup ($FileIni, $Group, $APP)
   Dim $Commands, $Command
   Dim $s, $i
   Dim $Name
   Dim $DestDir, $DestDirWork, $SubDirAPP, $DestSubDir
   Dim $IconIndex
   Dim $Minimize, $Replace, $RunInOwnSpace
'begin 
   $Commands = ReadSection($FileIni, $Group)
   $PathAPP = GetFilePath($FileIni)
   $i = 0
   for each $Command in $Commands
      $i = $i + 1
      $Name = $Command
      $s = ReadString ($FileIni, $Group, $Name, "")+"|"

      if blnDebug 
         LogAdd($Log,$LogFile,"I", $Group+"="+$s) 
      end if

      $CommandFile   = ExtractWord(1, $s, "|")
      $Arg           = ExtractWord(2, $s, "|")
      $WorkDir       = ExtractWord(3, $s, "|")
      $IconFile      = ExtractWord(4, $s, "|")
      $IconIndex     = 0
      $DestDir       = ExtractWord(5, $s, "|")
      $SubDirAPP     = ExtractWord(6, $s, "|")
      $WindowStyle   = ExtractWord(7, $s, "|")

      $DestDirWork   = GetDirWork($DestDir, $SubDirAPP)

      $DestSubDir    = $DestDirWork

      if blnDebug 
         LogAdd($Log,$LogFile,"I", $DestDir+"|"+$SubDirAPP+"|"+$DestDirWork)
      end if

      if Exist ($CommandFile)
         if blnDebug 
            LogAdd($Log,$LogFile,"I", "   "+$Name)
         end if

         '----------------------------------------------------------------

         if (UCase($DestDir) = UCase($FPrograms)) and (not IsTerminalApplication)
         'if (UCase($DestDir) = UCase($FPrograms)) and (@WKSTA <> $TerminalServer)

            '-----------------------------------------------------
            'if $SubDirAPP <> ""
            '   $DestSubDir =  $sPrograms+"\"+$SubDirAPP
            'else
            '   $DestSubDir =  $sPrograms
            'end if
            '-----------------------------------------------------
            CreateOne ($CommandFile, $Arg, $WorkDir, $IconFile, $Name, $DestDirWork, $DestSubDir, $WindowStyle, $PathAPP)
         else
         '----------------------------------------------------------------

         if (UCase($DestDir) = UCase($FTasks)) and (IsTerminalApplication)
         'if (UCase($DestDir) = UCase($FTasks)) and (@WKSTA = $TerminalServer)

            '-----------------------------------------------------
            'if $SubDirAPP <> ""
            '   $DestSubDir =  $sTasks+"\"+$SubDirAPP
            'else
            '   $DestSubDir =  $sTasks
            'end if
            '-----------------------------------------------------
            CreateOne ($CommandFile, $Arg, $WorkDir, $IconFile, $Name, $DestDirWork, $DestSubDir, $WindowStyle, $PathAPP)
         else
         if (UCase($DestDir) <> UCase($FPrograms)) and (UCase($DestDir) <> UCase($FTasks))
            CreateOne ($CommandFile, $Arg, $WorkDir, $IconFile, $Name, $DestDirWork, $DestSubDir, $WindowStyle, $PathAPP)
         end if
         end if
         end if
         '----------------------------------------------------------------

      else
         if $CommandFile <> ""
            if UCase($APP) <> UCase("Terminal")
               LogAdd($Log,$LogFile,"I", "ERROR: "+$Name+" Файл "+$CommandFile+" не существует или нет доступа")
            end if
         end if

         ' Delete Lnk

         if UCase($DestDir) = UCase($FPrograms) and (not IsTerminalApplication)
         'if UCase($DestDir) = UCase($FPrograms) and (@WKSTA <> $TerminalServer)

            if $SubDirAPP <> ""
               $DestSubDir =  $sPrograms+"\"+$SubDirAPP
            else
               $DestSubDir =  $sPrograms
            end if
            DeleteOne ($Name, $DestDirWork, $DestSubDir)
         else

         if UCase($DestDir) = UCase($FTasks) and (IsTerminalApplication)
         'if UCase($DestDir) = UCase($FTasks) and (@WKSTA = $TerminalServer)

            if $SubDirAPP <> ""
               $DestSubDir =  $sTasks+"\"+$SubDirAPP
            else
               $DestSubDir =  $sTasks
            end if
            DeleteOne ($Name, $DestDirWork, $DestSubDir)
         end if
         end if
      end if
   next
End Function

'------------------------------------------------------
' CreateLinksAPP ($FileIni, $Domain, $Group)
'------------------------------------------------------
function CreateLinksAPP ($FileIni, $Domain, $Group)
   Dim $FileINIApp, $LGroup, $i, $j, $n, $Done, $Object, $Objects, $s, $LGroups, $Item
'begin

   ' $FileINIApp = ReadString ($FileINI, $Domain, $Group, "")+"\"+strFileNameINI
   $FileINIApp = ReadString ($FileINI, $Domain, $Group, "")

   if Exist($FileINIApp) 

      if blnDebug
         LogAdd($Log,$LogFile,"I","Файл "+$FileINIApp+" существует")
      end if

      $Objects = ReadSections($FileINIAPP)
      for each $Object in $Objects
         $j = 0
         if $Object <> ""
            if blnDebug
               LogAdd($Log,$LogFile,"I",$Object)
            end if

            $n = WordCount($Object, ",")
            if $n > 0

               '------------------------------------------------------
               ' реализация ИЛИ
               '------------------------------------------------------
               $Done = 0
               $i = 1
               while ($i <= $n) and ($Done = 0)
                  $Item = ExtractWord($i, $Object, ",")
                  if $Item 

                     '------------------------------------------
                     ' User без домена
                     '------------------------------------------
                     if WordCount($Item, "\") = 1
                        $Item = $DOMEN_Default+"\"+$Item+"\"
                     end if

                     $LDomen = ExtractWord(1, $Item, "\")
                     $s = ExtractWord(2, $Item, "\")

                     '------------------------------------------
                     ' User
                     '------------------------------------------
                     if WordCount($Item, "\") = 3
                        if blnDebug
                           LogAdd($Log,$LogFile,"I","USER="+$Item)
                           LogAdd($Log,$LogFile,"I",UCase($LDomen+"\"+$s)+"="+UCase($DomainUser+"\"+$UserID))
                        end if

                        '------------------------------------------------
                        'if UCase($s) = UCase($UserID)
                        '   $Done = 1
                        'end if
                        '------------------------------------------------

                        if UCase($LDomen+"\"+$s) = UCase($DomainUser+"\"+$UserID)
                           $Done = 1
                        end if

                     else
                        if blnDebug
                           LogAdd($Log,$LogFile,"I","GROUP="+$Item)
                        end if
                        '------------------------------------------
                        ' GROUP
                        '------------------------------------------

                        if (UCase($DomainUser) = "GU"    and UCase($LDomen) = "DSP") or
                           (UCase($DomainUser) = "DSP"   and UCase($LDomen) = "GU") or
                           (UCase($DomainUser) = "ULNSK" and UCase($LDomen) = "DSP") or
                           (UCase($DomainUser) = "ULNSK" and UCase($LDomen) = "GU") 
'                           if blnDebug
                              LogAdd($Log,$LogFile,"I","Нет проверки вхождения в группу "+$Item+". ("+
                                 $DomainUser+"="+$LDomen+")")
'                           end if
                        else
                           If InGroup ($Item)
                              if blnDebug
                                 LogAdd($Log,$LogFile,"I","*  "+$Item)
                              end if
                              $Done = 1
                           else
                              if IndexOf($FUserAllGroups, $Item) >= 0
                                 if blnDebug
                                    LogAdd($Log,$LogFile,"I","** "+$Item)
                                 end if
                                 $Done = 1
                              end if
                              if blnDebug
                                 LogAdd($Log,$LogFile,"I","************** "+$Item)
                              end if
                           end if
                        end if
                     end if
                  end if
                  $i = $i + 1
               loop

               '------------------------------------------------------
               if $Done = 1
                  $DomainAPP = ExtractWord(1, $Item, "\")
                  if blnDebug
                     LogAdd($Log,$LogFile,"I","DomainAPP="+UCase($DomainAPP)+" "+
                                              "DomainUser="+UCase($DomainUser)+" "+
                                              "DomainPC="+UCase($DomainPC))
                  end if
                  select
                     case UCase($DomainAPP) = UCase($DomainDSP)
                        if (UCase($DomainUser) = UCase($DomainDSP)) and (UCase($DomainPC) = UCase($DomainDSP))
                           $j = $j + 1         
                           CreateLinksGroup ($FileINIApp, $Object, $Group)
                        else

                        '-------------------------------------------------------------
                        if (UCase($DomainUser) = UCase($DomainDSP)) and (UCase($DomainPC) <> UCase($DomainDSP))
                           $j = $j + 1         
                           CreateLinksGroup ($FileINIApp, $Object, $Group)
                        end if
                        '-------------------------------------------------------------

                        end if
                     case 1
                        $j = $j + 1         
                        CreateLinksGroup ($FileINIApp, $Object, $Group)
                  End Select
               end if
               '------------------------------------------------------

            end if
         end if
      next
   else
      if blnDebug
         LogAdd($Log,$LogFile,"I","Файл "+$FileINIApp+" не существует или нет доступа")
      end if
   end if
End Function

'------------------------------------------------------
' CreateLinks0 ($FileIni)
'------------------------------------------------------
function CreateLinks0 ($FileIni)
   Dim $Domains, $Groups, $Group, $Domain, $i, $DomainGroup
'begin 
   $Groups = AllGroups(0)
   for each $DomainGroup in $Groups
      $i = WordCount($DomainGroup, "\")
      if $i > 1
         $Domain = ExtractWord(1, $DomainGroup, "\")
         $Group = ExtractWord(2, $DomainGroup, "\")
         if ValueExists ($FileIni, $Domain, $Group) = True
            $DomainAPP = $Domain
            CreateLinksAPP ($FileIni, $Domain, $Group)
         end if
      end if
   next
End Function

'------------------------------------------------------
' CreateLinks1 ($FileIni)
'------------------------------------------------------
function CreateLinks1 ($FileIni)
   Dim $Domains, $Groups, $Group
'begin 
   $Domains = ReadSections($FileINI)
   for each $Domain in $Domains
      $Groups = ReadSection($FileINI, $Domain)
      for each $Group in $Groups
         if InGroup($Domain+"\"+$Group)
            $DomainAPP = $Domain
            CreateLinksAPP ($FileIni, $Domain, $Group)
         end if
      next
   next
End Function

'------------------------------------------------------
' CreateLinks2 ($FileIni)
'------------------------------------------------------
function CreateLinks2 ($FileIni)
   Dim $Domains, $Groups
'begin 
   $Domains = GetDomains
   for each $Domain in $Domains
      $Groups = LocalGroups($Domain)
      for each $Group in $Groups
         $DomainAPP = $Domain
         CreateLinksAPP ($FileIni, $Domain, $Group)
      next
   next
End Function

