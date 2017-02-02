#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.2
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here


#include <guiconstants.au3>
#include <constants.au3>

$AppName = "Outlook Profile Switcher"
if WinExists($AppName) = "1" then Exit

Dim $LastProfile
Dim $Cnt,$Str

Opt("WinTitleMatchMode", 2)
Opt("TrayMenuMode",1)   ; Default tray menu items (Script Paused/Exit) will not be shown.

$OutlookWinCap = "- Microsoft Outlook"
$OutlookPath = regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE\","Path")
$Outlookcommand  = $OutlookPath & "outlook.exe /profile "
$MainKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Profiles\"

; ############################### Create Main Window
$WinMain = GuiCreate($AppName,235,70,@DesktopWidth-300-100,@DesktopHeight-70-100)
$List = GUICtrlCreateCombo("", 10,10,140,80, BitOR($LBS_STANDARD,$CBS_SORT))
GetProfiles()
GetDefaultProfile()
$SwitchButton = GuiCtrlCreateButton("Switch", 160,10, 65, 22)
$MainProfileDefault = GUICtrlCreateCheckbox ("Set as default profile", 10, 40, 130, 20)
$LabelStatus = GUICtrlCreateLabel  ("lorem ipsum", 150, 43, 220, 20)
$LabelAutoSwitch = GUICtrlCreateCheckbox  ("Auto Switch", 10, 60, 130, 20)


; ############################### Shows Main window
Sleep(2000)
SplashOff()
While 1
    $msg = GUIGetMsg(1)
    If $msg[0] = $SwitchButton Then DoSwitch()

	If $msg[0] = $GUI_EVENT_MINIMIZE Then GuiSetState(@SW_HIDE,$WinMain)
    If $msg[0] = $GUI_EVENT_CLOSE And $msg[1] = $WinMain  Then DoExit() ;ExitLoop
    ;Show Gui
	GuiSetState(@SW_SHOWNORMAL,$WinMain)
;Get Tray
$TrayMsg = TrayGetMsg()
    Select
        Case $TrayMsg = 0
            ContinueLoop
        Case $TrayMsg = $TRAY_EVENT_PRIMARYDOWN
            GuiSetState(@SW_SHOWNORMAL,$WinMain)
	EndSelect
Wend

;############################## Switch between Profiles
Func DoSwitch()
$PID = ProcessExists("OUTLOOK.EXE")
	If $PID > 0 Then
		$Outlookrunning = "1"
	Else
		$Outlookrunning = "0"
	EndIf
		if GuiCtrlRead($MainProfileDefault) = "1" Then
			RegWrite($MainKey,"DefaultProfile","REG_SZ",GuiCtrlRead($List))
		EndIf

		$a=$Outlookcommand & chr(34) & GuiCtrlRead($List) & chr(34)
		If $Outlookrunning = "1" Then
			if GuiCtrlRead($List) <> $LastProfile Then
				GUICtrlSetData($LabelStatus,"Wait, switching profile..")
				GuiSetState(@SW_DISABLE,$WinMain) ; Turn the window off
				OutlookClose()
				run($a)
				$LastProfile = GuiCtrlRead($List)
				GuiSetState(@SW_ENABLE,$WinMain); Turns the window back on
				GUICtrlSetData($LabelStatus,"")
			Else
				MsgBox(64,$AppName,"Profil bereits aktiv",10)
			endif
		Else
			run($a)
			$LastProfile = GuiCtrlRead($List)
		EndIf
EndFunc

;############################## Close Outlook
Func OutlookClose()
	Do
		$PID = ProcessExists("OUTLOOK.EXE")
		run("taskkill /pid "&$PID,"",@SW_HIDE)
		;ProcessClose($PID)
		ProcessWaitClose($PID)
	 Until $PID = 0
  EndFunc

;############################## Reads Profiles from Registry
Func GetProfiles()
Do
$Cnt = $Cnt + 1
$var = RegEnumKey($MainKey, $cnt)
$Str = $Var & "|" & $Str
Until $var = "";
GUICtrlSetData($List,StringMid ($Str,2,StringLen($Str)-1))
$profiles = StringSplit(StringMid ($Str,2,StringLen($Str)-1),"|")
GUICtrlSetData($List,$profiles[1])
EndFunc

;############################## Reads Default Profile from Registry
Func GetDefaultProfile ()
$DefaultKey = regread($MainKey,"DefaultProfile")
if StringLen($DefaultKey)>0 Then
	GUICtrlSetData($List,$DefaultKey)
EndIf
EndFunc

;############################## Exits application
Func DoExit ()
If Not IsDeclared("iMsgBoxAnswer") Then Dim $iMsgBoxAnswer
$iMsgBoxAnswer = MsgBox(270628,$AppName,"Do you really want to quit this program?")
Select
   Case $iMsgBoxAnswer = 6 ;Yes
	exit
   Case $iMsgBoxAnswer = 7 ;No

EndSelect
EndFunc
