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

;Search the client name and press enter. This will open their rate sheet 
; WINWAIT VERSION
^+1::
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
Send, %client_code%!ad
WinWaitActive ^%client_code%
Sleep 1000
Send, ^c
Sleep 1000
WinActivate 2023 Church Insurance Renewal
Send, {Tab 2}{Control down}{Shift down}{Right 2}{Shift up}{Control up}{Backspace}{Down}^v (%client_code%){Tab}{Enter}
WinWaitActive - Foxit Reader$
Pause ; comment or uncomment if the laptop is being slow 
Send, {Tab}%client_code%{Tab}^v!{F4}{Enter}
Pause ; comment or uncomment if the laptop is being slow 
Sleep 1000 ;check and see how fast this is at closing and saving the pdf
Send, !ps{Enter}!pd
Send, !bb{Tab}1/9/2023{Tab}8:00 AM{Enter}
Sleep 2000
WinActivate ^%client_code%
Send, !aa{Tab 2}f{Tab 2}spr{Enter 2}
WinWaitActive %client_code% R ;make a regex here that can get endorsement names too 
Send, ^fD&O!f{Escape}{Right}^c!{F4}
WinActivate 2023 Insurance Renewal 
if (A_Clipboard = "2,000,000" or A_Clipboard = "5,000,000" or A_Clipboard = "10,000,000"){
	Send, ^fThe Directors and Officers!f{Escape}Z{Control down}{Shift down}{Down}{Backspace 2} ;add logic to delete D&O paragraph 
}
Return
;The rest is the manual pasting of the email/name of the client	

; PAUSE VERSION - ******HEYDE-1 is a non-D&O client you can test clipboard logic on****** also test on MAINS02
^+0::
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
Send, %client_code%{Enter}!ad
Pause
Send, ^c
Sleep 200
Send, !aa
WinActivate 2023 Church Insurance Renewal
Send, {Tab 2}{Down}^v (%client_code%){Tab}{Enter}
WinWaitActive PDF Reader$
;Pause ; comment or uncomment if the laptop is being slow 
Sleep 500
Send, {Tab}%client_code%{Tab}^v!{F4}{Enter}
Sleep 100
WinActivate 2023 Church Insurance Renewal
;Pause ; comment or uncomment if the laptop is being slow 
Send, !ps{Enter}!pd
Send, !bb{Tab}2023-05-29{Tab}8:30 AM{Enter}{Tab}{Down 2}{Control down}{Right}{Control up}
Pause
WinActivate ^%client_code%
Sleep 100
Send, {Tab 2}f{Tab 2}spr{Enter 2}
Return
;Pause
;WinWaitActive (^%client_code%|^Endorsement|^Reissue|^Updated)
;Send, ^fD&O!n{Escape}{Right}^c!{F4}
;Sleep 1000 ;remove later 
;WinActivate 2023 Church Insurance Renewal
;Sleep 1000
;clipboard logic is not working, try troubleshooting later 
if (A_Clipboard = "2,000,000" or A_Clipboard = "5,000,000" or A_Clipboard = "10,000,000"){
	Send, ^fThe Directors and Officers!f{Escape}{Control down}{Shift down}{Down}{Backspace 2}
}
Return

;if the PDF opens too slow and is left blank use Ctrl + Shft + 8 and it will fix it. Make sure you do not paste emails prior to this, it relies on the clipboard
^+8:: 
setTitleMatchMode RegEx
WinActivate PDF Reader$
Send, %client_code%{Tab}^v!{F4}{Enter}
Return