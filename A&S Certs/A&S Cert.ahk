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
-The script has logic to remove a double space in between the dash and the second year (like "2022-1-1 -  2023-1-1")
Eventually if this script is used for all A&S certs, this would be no longer needed. Tested the script on a version of a 
cert that doesnt have the double space and it works fine, it just moves the view down the page when it cant find the -  2.
The same would be true for new certs created with new clients, those seem to be made default without the double space.
-The years for the replacement need updated each year. Considered using A_YYYY to make it year-independent but that would mean
that when making January and onward certs "the year before" in Nov/Dec, it would require a whitelist of those months to get the 
years accurate
*/

^+0:: ;this key is the exact same as the below key, just that it calls itself so that you don't have to press the button every time
recursive_key(){
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
WinActivate ^Account Locate
Send, %client_code%{Enter}!aa
Pause
Send, {Tab 2}d{Tab 2}SP - Acc{Enter}
Sleep 100
Send, {Enter}
Pause
WinActivate ^%client_code%
Sleep 100
Send, !ap
WinActivate SP - Acc
Send, ^h2024-{Tab}2025-{Enter}!r{Enter}{Right}{Left}{Backspace}3{Tab}{Right}{Left}{Backspace}4{Enter}!r{Enter}
Send, -  2{Tab}- 2{Enter}!r{Enter}{Escape} ;-  2{Tab}- 2{Enter}!r{Enter}    {Tab} !f*19!r{Escape}{PgUp}
Pause
Send, !{F4}{Enter 2}
WinWaitActive ^Attachment Update
Send, {Enter}{Right}{Backspace}4
Pause
Send, {Tab}{Enter}
Sleep 100
Send, !aa
Sleep 2000
Send, !{F4}
WinWaitActive ^Account Locate
recursive_key()
}
recursive_key()
Return

^+9::
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
Send, %client_code%{Enter}!aa
Pause
Send, {Tab 2}d{Tab 2}SP - Acc{Enter}
Sleep 100
Send, {Enter}
Pause
WinActivate ^%client_code%
Sleep 100
Send, !ap
WinActivate SP - Acc
Send, ^h2023-{Tab}2024-{Enter}!r{Enter}{Right}{Left}{Backspace}2{Tab}{Right}{Left}{Backspace}3{Enter}!r{Enter}
Send, -  2{Tab}- 2{Enter}!r{Enter}{Escape} ;-  2{Tab}- 2{Enter}!r{Enter}    {Tab} !f*19!r{Escape}{PgUp}
Pause
Send, !{F4}{Enter 2}
WinWaitActive ^Attachment Update
Send, {Enter}{Right}{Backspace}3
Pause
Send, {Tab}{Enter}
Sleep 100
Send, !aa
Sleep 4000
Send, !{F4}
Return 
