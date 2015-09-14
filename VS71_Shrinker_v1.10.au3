; ----------------------------------------------------------------------------
;
; AutoIt Version: 3.0
; Language:       English
; Platform:       Win9x / NT
; Author:         Code Junkie
;
; Script Function:
;	Slim Down Visual Studio 2003
;
; ----------------------------------------------------------------------------


; ----------------------------------------------------------------------------
; Set up our defaults
; ----------------------------------------------------------------------------

;AutoItSetOption("MustDeclareVars", 1)
;AutoItSetOption("MouseCoordMode", 0)
;AutoItSetOption("PixelCoordMode", 0)
;AutoItSetOption("RunErrorsFatal", 0)
;AutoItSetOption("TrayIconDebug", 1)
;AutoItSetOption("WinTitleMatchMode", 4)

; ----------------------------------------------------------------------------
; Script Start
; ----------------------------------------------------------------------------

$Target = Inputbox("VS.NET 2003 Target","Directory Containing VS.NET 2003 Target:","y:")
If $Target = 1 then Exit
$Source = Inputbox("VS.NET 2003 Source","Directory Containing VS.NET 2003 Source:","z:")
If $Source = 1 then Exit



If @compiled = 1 Then CommonFiles()
If @compiled = 1 Then Run($Target & "\Setup\setup.exe")

GetMissingFiles()

Func GetMissingFiles()
  $LastFile = ""
  $Done = 0
  Do
    Do
       $Wait = WinWait("Microsoft Visual Studio .NET Enterprise Architect 2003 Setup","Error 1308.Source file not found:")
    Until $Wait = 1
    If StringInStr(WinGetText("Microsoft Visual Studio .NET Enterprise Architect 2003 Setup","Error 1308.Source file not found:"),"Ignore") <> 0 Then
      $FileLoc = StringTrimRight(StringTrimLeft(WinGetText("Microsoft Visual Studio .NET Enterprise Architect 2003 Setup","Error 1308.Source file not found:"),StringLen($Target)+57),59)
    Else
      $FileLoc = StringTrimRight(StringTrimLeft(WinGetText("Microsoft Visual Studio .NET Enterprise Architect 2003 Setup","Error 1308.Source file not found:"),StringLen($Target)+49),59)
    EndIf    
    If $FileLoc =  $LastFile then Error(2)
    CopyFile($FileLoc)
    $Success = ControlSend("Microsoft Visual Studio .NET Enterprise Architect 2003 Setup","Error 1308.Source file not found:","Button2","r")
    If $Success = 0 Then Error(1)
  Until $Done = 1

EndFunc

Func CommonFiles()
  Dim $Folders[2]
    $Folders[0] = "CommonAppData"
    $Folders[1] = "Program Files\Common Files"
  For $r = 0 to UBound($Folders,1) - 1
    CopyFolder($Folders[$r])
  Next
  Dim $Files[3]
    $Files[0] = "setup.exe"
    $Files[1] = "setup.ini"
    $Files[2] = "vs_setup.msi"
  For $r = 0 to UBound($Files,1) - 1
    CopyFile($Files[$r])
  Next
EndFunc


Func CopyFolder($Folder)
  DirCopy($Source & "\" & $Folder, $Target & "\" & $Folder)
EndFunc

Func Error($ErrorCode)
   Select
    Case $ErrorCode = 1
        If Msgbox(53,"Error","ControlSend Failed") = 2 Then Exit
    Case $ErrorCode = 2
        If Msgbox(53,"Error","CopyFile Failed") = 2 Then Exit
    Case $ErrorCode = 3
        MsgBox(0, "", "No preceding case was true!")
EndSelect
EndFunc

Func CopyFile($File)
  FileCopy($Source & "\" & $File, $Target & "\" & $File)
EndFunc