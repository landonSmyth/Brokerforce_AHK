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

1::
Send, {Backspace}4
Return

5::
Send, {Right}{Backspace 2}^v
Return