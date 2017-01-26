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

Opt("WinTitleMatchMode", 2)
Opt("TrayMenuMode",1)   ; Default tray menu items (Script Paused/Exit) will not be shown.

$OutlookWinCap = "- Microsoft Outlook"
$OutlookPath = regread("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE\","Path")
$Outlookcommand  = $OutlookPath & "outlook.exe /profile "
$MainKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Outlook\Profiles\"

