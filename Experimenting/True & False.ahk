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
Building := false
Contents := true
if (Building and Contents){
	MsgBox They have building & content insurance
} 
else if (Building and not Contents){
	MsgBox They have building but no contents
}
else if (not Building and Contents){
	MsgBox They have contents and no building
}
Return
