#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

`::
Pause -1
Return

PgUp::
reload
Return


^+0::
;WinWaitActive (^%clientName% R|^Endorsement) ;make a regex here that can get endorsement names too 
A_Clipboard := ""
Send, ^fD&O!f
Sleep 100 ;This sleep must be here to allow Excel to find the cell before moving forward 
Send, {Escape}{Right}
Send, {F2}{Control Down}{Shift Down}{Left}{Shift Up}{Control Up}^c
Pause
if (A_Clipboard = ""){
    MsgBox Clipboard is empty, remove D&O
    }
else
    MsgBox client has D&O, do not remove %A_Clipboard%
Return

^+1::
setTitleMatchMode RegEx
client := "MISSI34"
;window := "NEWLI-1 - New life in Christ Church"
;window := "ROSEC-1SEP - Rose City Community (Church of God)"
WinGetActiveTitle, window
if(InStr(window, "(") != 0 or InStr(window, ")") != 0){
	window := SubStr(window, 1, 13)
}
else {
	MsgBox This is fine, no bracket in name
}
WinActivate Account Locate
Sleep 500
WinActivate ^%window%
Sleep 500
WinActivate ^%client% R 23
Return
