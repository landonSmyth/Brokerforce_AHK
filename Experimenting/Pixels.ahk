#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

`::
Pause -1
Return

PgUp::
reload
Return

^+0::
PixelGetColor, Epic_Colour, 835, 545
MsgBox %Epic_Colour%
Return