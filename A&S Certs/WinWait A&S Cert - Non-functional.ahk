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

;---------------Hotkey for AB CHUR Clients---------------
^+1:: 
Send, !aa
Sleep 5000
Send, {Tab 2}d{Tab 2}SP - Acc{Enter}
Sleep 1000
Send, {Enter}
Sleep 5500
Send, ^h2023-{Tab}2024-{Enter}!r{Enter}{Right}{Left}{Backspace}2{Tab}{Right}{Left}{Backspace}3{Enter}!r{Enter}
Send, -  2{Tab}- 2{Enter}!r{Enter}{Escape}
Sleep 500
Send, !{F4}
Sleep 3000
Send, {Enter}{Right}
Return 

;---------------Hotkey for DB CHUR Clients---------------
^+2:: 
setTitleMatchMode 1
InputBox, client_name, Client Code, Enter client code
Send, %client_name%{Enter}!aa
WinWaitActive %client_name%
Sleep 1000
Send, {Tab 2}d{Tab 2}SP - Acc{Enter}
Pause ;at this point you highlight/hover the correct file and press `
Send, {Enter}
WinWaitActive SP - Accident Certificate
Send, ^h2023-{Tab}2024-{Enter}!r{Enter}{Right}{Left}{Backspace}2{Tab}{Right}{Left}{Backspace}3{Enter}!r{Enter}
Send, -  2{Tab}- 2{Enter}!r{Enter}{Escape}
Sleep 200 
Send, !{F4}
Sleep 3000 ;check to see if a WinWait would work here 
Send, {Enter}{Right}
Return 

^+0:: 
setTitleMatchMode 1
InputBox, client_name, Client Code, Enter client code
Send, {F2}
WinWaitActive Account
Sleep 1000
Send, %client_name%{Enter}


