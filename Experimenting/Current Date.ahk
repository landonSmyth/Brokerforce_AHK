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


^+1::
Send, %A_YYYY%-%A_MM%-%A_DD%
Sleep 20000
Return

;NOTE: Excel does not Tab out 05 12 as the date, it must be something
;like May 12*Tab*. Make a dictionary to assign the month numbers to the month
;string and send that in the main script instead of the month number