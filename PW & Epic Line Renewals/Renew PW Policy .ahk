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
-If a client does not have D&O, the Ctrl+F for "Non-Profit Org" will fail, and you just close the "Move to New Schedule" prompt
& reload the script 

-IF FRENCH POLICY: "Administrateurs et Dirigeants" is the wording the new schedule should be 
The line to move is 'Gestion du passif d'organismes sans but lucratif - par acte répréhensible' (form number CL5001E01)
*/


^+0::
recursive_key(){
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
Send, %client_code%{Enter}!ap
Pause ;should select one above the CHUR line that has the premium, if not, select it 
Send, {Down}+{F2}
Sleep 200
Send, {Tab 2}{Down 7}{Enter}
Pause
Send, !prr
Sleep 500 ; decrease as needed
Send, !lr
Pause ;at this point if the client is a new policy, have to create new policy and paste P04 number
Send, {Enter}
Pause ;variable loading time, wait until new policy is generated
Send, ^f
Sleep 500
Send, Non-Profit Org!f
Sleep 600
Send, {Escape}
Sleep 500
Send, !emn
Sleep 400
Send, Directors & Officers Liability
Pause ;delete later when functional
Send, {Enter}
Sleep 1000
Send, {Up 5}
WinWaitActive Account Locate
recursive_key()
}
recursive_key()
Return

RShift & RCtrl:: ;key to only update the forest fire limit to $500k
SetKeyDelay, 50 ;this must be here to allow policy works to keep up with input
Send, ^f
Sleep 500
Send, Forest Fire Fighting!f
Sleep 400
Send, {Enter}
Sleep 500
Send, {Escape}
Sleep 400
Send, {Right 4}
Sleep 200
SendMode Event
Send, 500000{Enter}
SendMode Input
Sleep 200
Send, {Down 5}{Left 4}
Return

RShift:: ;select the first of the four lines to be updated and then press hotkey
SetKeyDelay, 50 ;this must be here to allow policy works to keep up with input
Send, !e
Sleep 500
Send, e
Sleep 300
Send, !n
Sleep 200
Send, {Enter}
Sleep 300
Send, {Down}
Sleep 200
Send, !ee
Sleep 300
Send, !n
Sleep 200
Send, {Enter}
Sleep 300
Send, {Down}
Sleep 200
Send, !ee
Sleep 300
Send, !n
Sleep 200
Send, {Enter}
Sleep 300
Send, {Down}
Sleep 200
Send, !ee
Sleep 300
Send, !n
Sleep 200
Send, {Enter}
; Sleep 300
; Send, ^f
; Sleep 500
; Send, Forest Fire Fighting!f
; Sleep 400
; Send, {Enter}
; Sleep 500
; Send, {Escape}
; Sleep 400
; Send, {Right 4}
; Sleep 200
; SendMode Event
; Send, 500000{Enter}
; SendMode Input
; Sleep 200
; Send, {Down 5}{Left 4}
Return

PgDn::
Send, !{F4}
Sleep 350
Send, {Left}
Sleep 100
Send, {Enter}
Pause
Send, !{F4}
Return

^+8::
Send, Directors & Officers Liability
Return

^+9::
Send, !prr
Sleep 500 ; decrease as needed
Send, !lr
Pause ;at this point if the client is a new policy, have to create new policy and paste P04 number
Send, {Enter}
Pause ;variable loading time, wait until new policy is generated
Send, ^f
Sleep 500
Send, Legal Liability for Damage!f
Sleep 600
Send, {Escape}
; Sleep 500
; Send, !emn
; Sleep 400
; Send, Directors & Officers Liability
; Pause ;delete later when functional
; Send, {Enter}
; Sleep 1000
; Send, {Up 5}
Return