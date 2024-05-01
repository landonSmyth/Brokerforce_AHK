#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

PgUp & PgDn::Suspend ;used to temporarily turn off all hotkeys in the script if buttons needed for something else

PgUp::
reload
Return

`::
Send, Thank you for returning the completed renewal questionnaire and financials
Send, , I have added the documents to your file. 
Send, {Down 2}{Left 2}
Send, {Control down}{Shift down}{Left 2}{Shift up}{Control up}Have a great day