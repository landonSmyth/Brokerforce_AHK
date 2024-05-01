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
-If Microsoft Outlook ever moves the Ctrl+F function to something other than forward email, it won't work and would need updated
-The tab counts in Epic might not stay the same forever, may need updated in the future
-This only works for CHUR clients, CAPSS questionnaire reminders cannot be sent with this workflow
-The 3rd reminder template indicates that a copy of last year is attached so if it is a client's first time doing the questionnaire,
this line would have to be manually removed.
-The winactivate logic for closing the initial renewal questionnaire email relies on the front of the subject line
to be "202# Church Insurance Renewal Questionnaire". If it is something different it will close the wrong window
-The program forwards the renewal questionnaire email to any and all emails contained in the top recipient list.
If there are any recipients that are in the CC line, those emails would need copied and pasted into that field manually 
-The logic for automatically storing the names of the recipients relies on the greeting being "Hello *name*," and the 
first word on the following line to be "This". These conditions are met if reminders are sent forwarding an email that 
was originally sent with one of the templates in this script. If sending a reminder that does not meet these conditions,
the name of the recipient will be incorrect and will need manually typed out
*/
;---------------Before Running: Open the client by clicking on the activity in Epic Home---------------


;---------------1st Reminder---------------
^+1::
;---------------Email Template Variables---------------

Abuse_D_O := "
(
This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement for your file. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Additionally, as you have abuse liability coverage you are required to complete the abuse application section this year (Questions 16-23 on the questionnaire).

The insurance company is doing a review and needs to make sure that their clients with abuse coverage are following all requirements to continue qualifying for an increased abuse limit.

Questions 16 – 22 on the application are the requirements that need to be included in your protocol and followed. To confirm that they are being followed, we will need you to return the form with the responses answered “yes” (excluding question 17 items 2&3).

Please let me know if you have any questions. 
)"

Abuse_Only := "
(
This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire and return it to me. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Additionally, as you have abuse liability coverage you are required to complete the abuse application section this year (Questions 16-23 on the questionnaire).

The insurance company is doing a review and needs to make sure that their clients with abuse coverage are following all requirements to continue qualifying for an increased abuse limit.

Questions 16 – 22 on the application are the requirements that need to be included in your protocol and followed. To confirm that they are being followed, we will need you to return the form with the responses answered “yes” (excluding question 17 items 2&3).

Please let me know if you have any questions. 
)"

D_O_Only := "
(
This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement for your file. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Please let me know if you have any questions. 
)"

Neither_coverage := "
(
This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire and return it to me. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Please let me know if you have any questions. 
)"

;---------------Follow up date variable, set this before running script---------------
follow_up := "2024-04-29"

setTitleMatchMode RegEx
WinGetActiveTitle, client_window
month_to_number := {"JAN":1, "FEB":2, "MAR":3, "APR":4, "MAY":5, "JUN":6, "JUL":7, "AUG":8, "SEP":9, "OCT":10, "NOV":11, "DEC":12}
month := SubStr(client_window, 8, 3)
month_number := month_to_number[month]

Send, !aa
Pause
Send, {Tab 2}d{Tab 2}%month_number% Ren Q Sent{Enter 2}
WinWaitActive .*Church Insurance Renewal Questionnaire for.*
Sleep 100
Send, ^a
Sleep 50
Send, ^c
Sleep 100
left_trim := RegExMatch(A_Clipboard, "Dear ") ;this needs trimming when you end up using it because the hard coded numbers are for "Hello "
right_trim := RegExMatch(A_Clipboard, "As")
recipient_names := SubStr(A_Clipboard, left_trim+5, (right_trim - left_trim - 11))
Send, {Tab 5}
Sleep 100
Send, ^a^c
Sleep 200
Send, ^f
Sleep 500 ;decrease if possible
Send, ^v{Tab 2}{Control Down}{Shift Down}{Right}{Shift Up}{Control Up}
Send, Reminder{Tab 2}
Sleep 100
right_cut := RegExMatch(client_window, " - ")
client_code := SubStr(client_window, 1, right_cut-1)
WinActivate ^%client_code%
Sleep 100
Send, {Tab 5}{Enter}
Sleep 1000
Send, !c{Tab}f{Tab 2}spr{Enter 2}
;Pause
WinWaitActive Excel$
Sleep 1250 ;this likely would need tuning depending on the system or the loading time of Excel
A_Clipboard := ""
Send, ^fAbuse!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 100
if (A_Clipboard = ""){
	Abuse := false
} else {
	Abuse := true
}
A_Clipboard := ""
Send, ^fD&O!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 100
if (A_Clipboard = ""){
	D_O := false
} else {
	D_O := true
}
Send, !{F4}{Right}{Enter}
Sleep 200
WinActivate ^%client_code%
Sleep 100
Send, {Tab 5}{Enter}
WinActivate ^\d{4} Church Insurance Renewal Questionnaire ;this is supposed to activate the initial email and close it
Sleep 100
Send, !{F4}
Sleep 100
WinActivate Reminder
if (Abuse and D_O){
	Send, Hello %recipient_names%,{Enter 2}%Abuse_D_O%
	Send, ^fPlease complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement for your file
	Send, !f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (Abuse and not D_O){
	Send, Hello %recipient_names%,{Enter 2}%Abuse_Only%
	Send, ^fPlease complete the attached questionnaire and return it to me.!f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (not Abuse and D_O){
	Send, Hello %recipient_names%,{Enter 2}%D_O_Only%
	Send, ^fPlease complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement for your file
	Send, !f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (not Abuse and not D_O){
	Send, Hello %recipient_names%,{Enter 2}%Neither_Coverage%
	Send, ^fPlease complete the attached questionnaire and return it to me.!f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
;here you review and make any necessary changes such as attaching last year for a slow client, then send the email and drag over to attach
WinWaitActive Attach To
Sleep 300
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}{Tab 2}%month_number% Ren Q 1st Reminder Sent{Tab 2}q{Tab}
Pause
Send, {Enter}
Sleep 150
setTitleMatchMode 1
while WinExist("Attach To"){
	Sleep 200
}
setTitleMatchMode RegEx
Send, !av
Click, 264 679
WinWaitActive Add a Note
Sleep 100
Send, 1st reminder sent
Pause ;here you would add any additional notes needed
Send, {Tab}{Enter}
Sleep 200
Send, !d
Sleep 50
Send, %follow_up%
Sleep 300
Send, ^s
Pause
Send, !{F4}
Return


;---------------2nd Reminder---------------
^+2::
;---------------Email Template Variables---------------

Abuse_D_O := "
(
This is another reminder to let you know we have not received your renewal questionnaire for this year. 

Please complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement as soon as possible. If the financials are not yet available, please send the questionnaire separately for the time being and let me know when the financials are expected to be ready/finalized.

The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Additionally, as you have abuse liability coverage you are required to complete the abuse application section this year (Questions 16-23 on the questionnaire).

The insurance company is doing a review and needs to make sure that their clients with abuse coverage are following all requirements to continue qualifying for an increased abuse limit.

Questions 16 – 22 on the application are the requirements that need to be included in your protocol and followed. To confirm that they are being followed, we will need you to return the form with the responses answered “yes” (excluding question 17 items 2&3). 

Please let me know if you have any questions.
)"

Abuse_Only := "
(
This is another reminder to let you know we have not received your renewal questionnaire for this year. 

Please complete the attached questionnaire and return it to me as soon as possible.

The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Additionally, as you have abuse liability coverage you are required to complete the abuse application section this year (Questions 16-23 on the questionnaire).

The insurance company is doing a review and needs to make sure that their clients with abuse coverage are following all requirements to continue qualifying for an increased abuse limit.

Questions 16 – 22 on the application are the requirements that need to be included in your protocol and followed. To confirm that they are being followed, we will need you to return the form with the responses answered “yes” (excluding question 17 items 2&3). 

Please let me know if you have any questions.
)"

D_O_Only := "
(
This is another reminder to let you know we have not received your renewal questionnaire for this year. 

Please complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement as soon as possible. If the financials are not yet available, please send the questionnaire separately for the time being and let me know when the financials are expected to be ready/finalized.

The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file.

Please let me know if you have any questions.
)"

Neither_coverage := "
(
This is another reminder to let you know we have not received your renewal questionnaire for this year. 

Please complete the attached questionnaire and return it to me as soon as possible.

The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file.

Please let me know if you have any questions.
)"

;---------------Follow up date variable, set this before running script---------------
follow_up := "2024-04-15"

setTitleMatchMode RegEx
WinGetActiveTitle, client_window
month_to_number := {"JAN":1, "FEB":2, "MAR":3, "APR":4, "MAY":5, "JUN":6, "JUL":7, "AUG":8, "SEP":9, "OCT":10, "NOV":11, "DEC":12}
month := SubStr(client_window, 8, 3)
month_number := month_to_number[month]

Send, !aa
Pause
Send, {Tab 2}d{Tab 2}Reminder Sent{Enter 2}
WinWaitActive .*Church Insurance Renewal Questionnaire for.*
Sleep 100
Send, ^a
Sleep 50
Send, ^c
Sleep 100
left_trim := RegExMatch(A_Clipboard, "Hello ")
right_trim := RegExMatch(A_Clipboard, "This")
recipient_names := SubStr(A_Clipboard, left_trim+6, (right_trim - left_trim - 11))
if (RegExMatch(A_Clipboard, "financial statement") != 0){
	D_O := True
}
else{
	D_O := False
}
if (RegExMatch(A_Clipboard, "abuse liability coverage") != 0){
	Abuse := True
}
else{
	Abuse := False
}
Send, {Tab 5}
Sleep 100
Send, ^a^c
Sleep 200
Send, ^f
Sleep 500 ;decrease if possible
Send, ^v{Tab 2}{Control Down}{Shift Down}{Right}{Shift Up}{Control Up}+{Right}
Send, 2nd{Tab 2}
Sleep 100
right_cut := RegExMatch(client_window, " - ")
client_code := SubStr(client_window, 1, right_cut-1)
if (Abuse and D_O){
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%Abuse_D_O%
	Send, ^fPlease complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement as soon as possible
	Send, !f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (Abuse and not D_O){
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%Abuse_Only%
	Send, ^fPlease complete the attached questionnaire and return it to me as soon as possible.!f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (not Abuse and D_O){
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%D_O_Only%
	Send, ^fPlease complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement as soon as possible
	Send, !f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else if (not Abuse and not D_O){
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%Neither_Coverage%
	Send, ^fPlease complete the attached questionnaire and return it to me as soon as possible.!f{Escape}^u{Down}
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
while (True){ ;Once the reminder email is sent, close the previous reminder
	if not WinExist("^2nd Reminder.*Church Insurance Renewal Questionnaire"){ 
		if WinExist("^Reminder.*Church Insurance Renewal Questionnaire"){
			WinActivate ^Reminder.*Insurance Renewal Questionnaire ;this is supposed to activate the first reminder and close it
			Sleep 100
			Send, !{F4}
		}
		break
	}
	else{
		Sleep 500
	}
}
WinWaitActive Attach To
Sleep 150
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}{Tab 2}%month_number% Ren Q 2nd Reminder Sent{Tab 2}q{Tab}
Pause
Send, {Enter}
Sleep 150
setTitleMatchMode 1
while WinExist("Attach To"){
	Sleep 200
}
setTitleMatchMode RegEx
Send, !av
Sleep 500
Click, 264 679
WinWaitActive Add a Note
Sleep 100
Send, 2nd reminder sent
Pause ;here you would add any additional notes needed
Send, {Tab}{Enter}
Sleep 200
Send, !d
Sleep 50
Send, %follow_up%
Sleep 300
Send, ^s
Pause
Send, !{F4}
Return

;---------------3rd Reminder---------------
^+3::
;---------------Email Template Variables---------------
Yes_D_O := "
(
This is another reminder to let you know we have still not received your renewal questionnaire for this year. Please see attached the blank questionnaire as well as a copy of the church’s questionnaire from last year for your reference. 

Please complete the attached questionnaire and return it to me along with a copy of the church’s most recent year-end financial statement AS SOON AS POSSIBLE. If the financials are not yet available, please send the questionnaire separately for the time being and let me know when the financials are expected to be ready/finalized.
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire must be present in your file. 
 
Your prompt action and response would be highly appreciated.
)"

No_D_O := "
(
This is another reminder to let you know we have still not received your renewal questionnaire for this year. Please see attached the blank questionnaire as well as a copy of the church’s questionnaire from last year for your reference. 

Please complete the attached questionnaire and return it to me AS SOON AS POSSIBLE. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire must be present in your file. 
 
Your prompt action and response would be highly appreciated.
)"

;---------------Follow up date variable, set this before running script---------------
follow_up := "2024-04-24"

setTitleMatchMode RegEx
WinGetActiveTitle, client_window
month_to_number := {"JAN":1, "FEB":2, "MAR":3, "APR":4, "MAY":5, "JUN":6, "JUL":7, "AUG":8, "SEP":9, "OCT":10, "NOV":11, "DEC":12}
month := SubStr(client_window, 8, 3)
month_number := month_to_number[month]

Send, !aa
Pause
Send, {Tab 2}d{Tab 2}Reminder Sent{Enter 2}
WinWaitActive .*Church Insurance Renewal Questionnaire for.*
Sleep 100
Send, ^a
Sleep 50
Send, ^c
Sleep 100
left_trim := RegExMatch(A_Clipboard, "Hello ")
right_trim := RegExMatch(A_Clipboard, "This")
recipient_names := SubStr(A_Clipboard, left_trim+6, (right_trim - left_trim - 11))
if (RegExMatch(A_Clipboard, "financial statement") != 0){
	D_O := True
}
else{
	D_O := False
}
Send, {Tab 5}
Sleep 100
Send, ^a^c
Sleep 200
Send, ^f
Sleep 500 ;decrease if possible
Send, ^v{Tab 2}{Control Down}{Shift Down}{Right 3}{Shift Up}{Control Up}
Send, 3rd {Tab 2}
right_cut := RegExMatch(client_window, " - ")
client_code := SubStr(client_window, 1, right_cut-1)
if (D_O){
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%Yes_D_O%
	Sleep 500
	Send, {Up 9}
	Send, ^fPlease complete the attached questionnaire and return it to me along with a copy of the church's most recent year-end financial statement
	Send, !f{Escape}^u{Right 2}{Control Down}{Shift Down}{Right 4}{Shift Up}{Control Up}^b^!h
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
else {
	Send, REMINDER:+{Up}^u^b{Left}{Down}{Enter 2}{Up}Hello %recipient_names%,{Enter 2}%No_D_O%
	Sleep 500
	Send, {Up 9}
	Send, ^fPlease complete the attached questionnaire and return it to me
	Send, !f{Escape}^u{Right 2}{Control Down}{Shift Down}{Right 4}{Shift Up}{Control Up}^b^!h
	Send, ^fThe renewal questionnaire is a requirement from the Insurance Co!f{Escape}^u^b{Up}
}
while (True){ ;Once the reminder email is sent, close the previous reminder
	if not WinExist("^3rd Reminder.*Church Insurance Renewal Questionnaire"){ 
		if WinExist("^2nd Reminder.*Church Insurance Renewal Questionnaire"){
			WinActivate ^2nd Reminder.*Insurance Renewal Questionnaire ;this is supposed to activate the second reminder and close it
			Sleep 100
			Send, !{F4}
		}
		break
	}
	else{
		Sleep 500
	}
}
WinWaitActive Attach To
Sleep 150
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}{Tab 2}%month_number% Ren Q 3rd Reminder Sent{Tab 2}q{Tab}
Pause
Send, {Enter}
Sleep 150
setTitleMatchMode 1
while WinExist("Attach To"){
	Sleep 200
}
setTitleMatchMode RegEx
Send, !av
Click, 264 679
WinWaitActive Add a Note
Sleep 100
Send, 3rd reminder sent
Pause ;here you would add any additional notes needed
Send, {Tab}{Enter}
Sleep 200
Send, !d
Sleep 50
Send, %follow_up%
Sleep 300
Send, ^s
Pause
Send, !{F4}
Return


;---------------Key to write email bit on copy of last year included---------------
^+5::
Send, {Control Down}{Up 14}{Control Up}
Sleep 100
Send, ^fPlease complete the!f{Escape}{Left 3}
Send, {Space}Please see attached the blank questionnaire as well as a copy of the church’s questionnaire from last year for your reference.
Return

^+6::
setTitleMatchMode RegEx
while (True){ ;if the reminder email is sent, close the previous reminder
	if not WinExist("^2nd Reminder.*Church Insurance Renewal Questionnaire"){ 
		if WinExist("^Reminder.*Church Insurance Renewal Questionnaire"){
			WinActivate ^Reminder.*Insurance Renewal Questionnaire ;this is supposed to activate the first reminder and close it
			Sleep 100
			Send, !{F4}
		}
		break
	}
	else{
		Sleep 500
	}
}
Return

;testing
^+7::
Send, ^a
Sleep 50
Send, ^c
Sleep 100
left_trim := RegExMatch(A_Clipboard, "Hello ")
right_trim := RegExMatch(A_Clipboard, "As")
MsgBox left trim is %left_trim%`nright trim is %right_trim%
if (RegExMatch(A_Clipboard, "financial statement") != 0){
	D_O := True
}
else{
	D_O := False
}
if D_O
	MsgBox they have d and o coverage 
recipients := SubStr(A_Clipboard, left_trim+6, (right_trim - left_trim - 11))
MsgBox %recipients%
Return

;address_untrimmed := A_Clipboard
;left_trim := RegExMatch(address_untrimmed, "Address:")
;right_trim := RegExMatch(address_untrimmed, "City:")
;address := SubStr(address_untrimmed, left_trim+10, right_trim-15)





