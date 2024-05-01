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

^+0::
recursive_key(){
setTitleMatchMode RegEx
WinWaitActive ^\d{4} Church Insurance Renewal
Sleep 100 
Send {LAlt Down}pdbb{LAlt Up}{Tab}2024-03-25{Tab}8:00 AM
Sleep 700 ;decrease once comfortable
Send, {Enter}
Sleep 50
Pause
Send, !s
Sleep 1000
recursive_key()
}
recursive_key()
Return