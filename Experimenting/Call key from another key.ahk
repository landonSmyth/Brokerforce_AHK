#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

PgUp & PgDn::Suspend ;used to temporarily turn off all hotkeys in the script if buttons needed for something else

`::
Pause -1
Return

PgUp::
reload
Return

activate_ahk_windows(){
	setTitleMatchMode RegEx
	while (True){
		IfWinExist Renewal premium
		{
			Sleep 200
			WinActivate Renewal premium
			Sleep 2000
		}
		IfWinExist ^Check Within 6
		{
			Sleep 200
			WinActivate ^Check Within 6
		}
		else {
			Sleep 100
		}
	}	
}

^+1::
Run C:\Users\LandonSmyth\OneDrive - Brokerforce Insurance Inc\Documents\AutoHotkey\Experimenting\Activate ahk Windows.ahk
InputBox, premium, Renewal premium, enter the renewal premium
MsgBox, 0, Check Within 6`% Increase, check six percent increase
Return

^+2::
activate_ahk_windows()
Return

