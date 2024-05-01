#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

/*
Drag over email from outlook into client and select the policy 
Make sure that the cursor is in the policy field and the correct 
Policy has been selected (as in press enter or double click the policy)
*/

Del::ExitApp

PgUp & PgDn::Suspend ;used to temporarily turn off all hotkeys in the script if buttons needed for something else

`::
Pause -1
Return

PgUp::
reload
Return

^+0:: ;activates the key itself but requires the explicit month number declaration
recursive_key(){
setTitleMatchMode RegEx
;---------------Set these two variables before running script---------------

month := "10"
follow_up := "2024-07-16"

;---------------------------------------------------------------------------
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
WinActivate Account Locate
Send, %client_code%{Enter}!aa
WinWaitActive ^Attach 
Sleep 200
Send, {F4}
Sleep 300 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}{Tab 2}%month% Ren Q Sent{Tab 2}q{Tab}
Pause
Send, {Enter}
Sleep 150
Send, !av
Sleep 100
Send, ^n{Enter}mi{Tab 2}%month% Ren Q Sent{Tab 2}smy{Tab 2}%follow_up%{Enter}
Pause
Send, !{F4}
WinActivate ^Delayed
Sleep 200
Send, !hztum
Sleep 500
Send, {Up}
WinActivate Account Locate
recursive_key()
}
recursive_key()
Return


month := "11"
follow_up := ""

setTitleMatchMode RegEx
WinWaitActive ^Attach 
Sleep 300
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}{Tab 2}%month% Ren Q Sent{Tab 2}q{Tab}
Pause
Send, {Enter}
Sleep 150
Send, !av
Sleep 100
Send, ^n{Enter}mi{Tab 2}%month% Ren Q Sent{Tab 2}smy{Tab 2}2023-09-18{Enter}
Pause
Send, !{F4}
WinActivate ^Delayed
Sleep 200
Send, !hztum
Sleep 500
Send, {Up}
WinActivate Account Locate
Return

^+1::
Sleep 100
Send, ^n{Enter}mi{Tab 2}9 Ren Q Sent{Tab 2}smy{Tab 2}2023-07-17{Enter}
Return


^+3:: ; copy of broken epic integration version
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
WinActivate Account Locate
Send, %client_code%{Enter}!av
Pause
WinActivate ^Delayed
Send !hy1
Pause
Send, {Tab}%client_code%{Tab}{Space}{Tab}p{Tab}{F4}
Sleep 200
Send, {Tab}{Enter}{Tab 2}9 Ren Q Sent{Tab 2}q
Pause ;check everything is working and remove when functional
Send, {Enter}
Pause
WinActivate ^%client_code%
Sleep 100
Send, ^n{Enter}mi{Tab 2}9 Ren Q Sent{Tab 2}smy{Tab 2}2023-07-17{Enter}
;if you ever give the script to anyone else, what follows might not be preferred
WinActivate ^Delayed
Sleep 200
Send, !hztum
Return

/*
the number in the Ren Q sent and date need to be edited with respect 
to the month in which you are sending the questionnaires 
*/