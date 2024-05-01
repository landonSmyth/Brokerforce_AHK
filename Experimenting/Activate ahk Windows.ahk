#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

ahk_windows_on_top(){
	setTitleMatchMode RegEx
	while (True){
		IfWinExist Renewal premium
		{
			Sleep 200
			WinActivate Renewal premium
			;Sleep 2000
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
ahk_windows_on_top()