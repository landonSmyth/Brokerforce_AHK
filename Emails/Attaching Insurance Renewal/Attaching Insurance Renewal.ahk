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


/*
---------------Notes and Dependencies---------------
-the winactivate for un-flagging the email in outlook relies on the folder being named "Delayed Renewals"
*/

^+0:: ;faster key, hard coded year but requires no extra sleep time for clipboard 
recursive_key(){

renewal_year := "2023"

setTitleMatchMode RegEx
WinWaitActive ^Attach To
Sleep 150
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not an auto or event policy
Send, {Tab}{Enter}{Tab 2}2023 Insurance Renewal{Tab 2}pol{Tab}
Pause
Send, {Enter}
Pause
Send, !{F4}
WinActivate ^Delayed Ren
Sleep 200
Send, !hztum
Sleep 500
Send, {Up}
Sleep 20
WinActivate Account Locate
recursive_key()
}
recursive_key()
Return

^+9:: ;reliable key, year independent but requires an extra sleep
recursiveKey(){
setTitleMatchMode RegEx
WinWaitActive ^Attach To	
Sleep 150
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not an auto or event policy
Send, {Tab}{Enter}{Tab 2}
Sleep 50
Send, ^c
Sleep 100
renewal_year := SubStr(A_Clipboard, 1, 4)
Send, %renewal_year% Insurance Renewal{Tab 2}pol{Tab}
Pause
Send, {Enter}
WinWaitActive ^Account Locate
recursiveKey()
}
recursiveKey()
Return

