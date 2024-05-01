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
Client_Codes := ["BIBLI-1", "BLACK03", "BLESS-2"]
For index, element in Client_Codes
	Sleep 1000
	Send, %element%
Return


^+1::
MonthArray := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
month := 7
nextMonth := MonthArray[month + 1]
MsgBox, %nextMonth%
Return