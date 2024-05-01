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

;---------------Hotkey for clients that do not have A&S---------------
^+0:: 
recursive_key(){
setTitleMatchMode RegEx
InputBox, coverage, Accident and Sickness?, Enter 1 if the client does not have A and S. Enter 2 if they do
if (coverage = 1){
    InputBox, client_code, Client Code, Enter client code
    StringUpper, client_code, client_code
    Send, %client_code%{Enter}!ap
    Pause
    Send, !tr
    Sleep 100
    Send, {Tab 8}{Backspace}{Tab}{Backspace}{Tab 3}REN
    Pause
    Send, {Enter}
    Pause
    Send, !{F4}
    WinWaitActive ^Account Locate
    }
else if (coverage = 2){
    InputBox, client_code, Client Code, Enter client code
    StringUpper, client_code, client_code
    Send, %client_code%{Enter}!ap
    Pause
    Send, !tr
    Sleep 100
    Send, {Tab 8}{Backspace}{Tab}{Backspace}{Tab 3}REN
    Pause
    Send, {Enter}
    Sleep 500
    Send, {Down 2}
    Sleep 1000
    Send, !tr
    Pause
    Send, {Enter}
    Pause
    Send, !{F4}
    WinWaitActive ^Account Locate
    }
recursive_key()
}
recursive_key()
Return

;---------------Hotkey for clients that have A&S---------------
^+9:: 
recursive_key2(){
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
Send, %client_code%{Enter}!ap
Pause
Send, !tr
Sleep 100
Send, {Tab 8}{Backspace}{Tab}{Backspace}{Tab 3}REN
Pause
Send, {Enter}
Sleep 500
Send, {Down 2}
Sleep 1000
Send, !tr
Pause
Send, {Enter}
Pause
Send, !{F4}
WinWaitActive ^Account Locate
recursive_key2()
}
recursive_key2()
Return
