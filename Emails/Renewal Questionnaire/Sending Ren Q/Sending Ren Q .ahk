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

; Pause/Reliable Version
^+0::
recursive_key(){
;---------------Variables that define delay-send date, update as needed---------------
send_date := "2024-04-24"
send_time := "8:00 AM"

setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
WinActivate ^Account Locate
Send, %client_code%{Enter}!aa
WinWaitActive ^%client_code%
WinGetActiveTitle, client_window
client_name := SubStr(client_window, 14)
if(InStr(client_window, "(") != 0 or InStr(client_window, ")") != 0){
	client_window := SubStr(client_window, 1, 13)
}
Sleep 100
WinActivate .*Church Insurance Renewal Questionnaire.*
Send, {Tab 2}{Down}%client_name% (%client_code%){Tab}{Enter}
Sleep 500
WinActivate .*Church Insurance Renewal Questionnaire.*
Sleep 50
Send, !ps{Enter}!pd
Sleep 20
Send, !bb{Tab}%send_date%{Tab}%send_time%{Enter}{Tab}{Down 2}
Sleep 650
WinActivate ^Renewal Questionnaire.*PDF Reader
Sleep 100
Send, {Tab}%client_code%{Tab}%client_name%
Sleep 400
WinActivate ^%client_window%
Sleep 100
WinActivate ^%client_window%
Sleep 200
Send, {Tab 2}f{Tab 2}spr{Enter 2}
WinWaitActive Excel$
Sleep 1000 ;this likely would need tuning depending on the system or the loading time of Excel
Send, !hzefdg
Sleep 500
Send, a48{Enter}
Sleep 100
A_Clipboard := ""
Send, ^fD&O!f
Sleep 200 ;This sleep must be here to allow Excel to find the cell before moving forward 
Send, {Escape}{Right}
Sleep 200 ;this is here so you can get a glance and make sure that it is actually working correctly
Send, {F2}{Control Down}{Shift Down}{Left}{Shift Up}{Control Up}^c
Sleep 100
Send, !{F4}{Right}{Enter}
Sleep 200
WinActivate ^Renewal Questionnaire.*PDF Reader
Sleep 50
Send, !{F4}{Enter}
Sleep 250
WinActivate .*Church Insurance Renewal Questionnaire.*
While, DllCall("GetOpenClipboardWindow")
	Sleep, 10
if (A_Clipboard = ""){
	Send, ^fThe Directors and Officers!f{Escape}{Control down}{Shift down}{Down}{Shift Up}{Control Up}{Backspace 2}
}
Sleep 100

;begin auto email pasting section
WinActivate ^%client_window%
Sleep 100
Click, 121 592  
Sleep 100
Click, 121 592 Right 
Sleep 100
Click 185 629 ;update this
Sleep 100
Click, 121 592 Right 
Sleep 100
Click, 174 604
Sleep 100
info := A_Clipboard
left_trim := RegExMatch(info, "-\d{4}")
right_trim := RegExMatch(info, "Contact via")
emails := SubStr(info, left_trim+8, -20)
A_Clipboard := emails
WinActivate .*Church Insurance Renewal Questionnaire.*
Send, !s{Enter}
Sleep 50
Send, ^v ;%emails%
Sleep 350
Send, {Enter}
;Pause ;here you click on the email to make it select the contact
Send, {Tab 4}
Sleep 100
;end auto email pasting section 

Send, ^fDear!f{Escape}{Right 2}
A_Clipboard := client_name
setTitleMatchMode 1
Sleep 500
WinGetActiveTitle, email
while WinExist(email){
	Sleep 200
}
setTitleMatchMode RegEx
WinActivate ^%client_window%
Send, !{F4}
WinActivate ^Account Locate
;WinWaitActive ^Account Locate
recursive_key()
}
recursive_key()
Return

;if the PDF opens too slow and is left blank use Ctrl + Shft + 8 and it will fix it. Make sure you do not paste emails prior to this, it relies on the clipboard
^+8:: 
setTitleMatchMode RegEx
WinActivate PDF Reader$
Send, %client_code%{Tab}^v!{F4}{Enter}
Return

^+3:: ;testing for D&O logic - remove later
A_Clipboard := ""
Send, ^fD&O!f
Sleep 100 ;This sleep must be here to allow Excel to find the cell before moving forward 
Send, {Escape}{Right}
Send, {F2}{Control Down}{Shift Down}{Left}{Control Up}{Shift Up}^c
Sleep 100
Send, !{F4}{Right}{Enter}
Pause ;remove later once the logic is working 
WinActivate .*Church Insurance Renewal Questionnaire.*
if (A_Clipboard = ""){
	Send, ^fThe Directors and Officers!f{Escape}{Control down}{Shift down}{Down}{Shift Up}{Control Up}{Backspace 2}
}
Return

; WINWAIT VERSION
^+1::
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
Send, %client_code%!ad
WinWaitActive ^%client_code%
Sleep 1000
Send, ^c
Sleep 1000
WinActivate 2024 Church Insurance Renewal
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