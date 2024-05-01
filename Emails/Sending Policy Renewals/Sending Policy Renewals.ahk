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
-For the first time running the script for a session, make sure that the default printer in Microsoft Word when doing Alt+P+P
is print to pdf. Otherwise it will not work when you get to saving an A&S cert or renewal letter 
-The payment / renewal date replacement depends on the email template date body having "Month , 202#".
The spacing must be the same for the replacement logic to work.
-The email template must be named "AB - 202# Insurance Renewal ()*Spacebar*". AB - or DB - does not matter,
but the logic that fills out the email subject must have that as the subject name. This would really only be a problem if the name
was changed or if someone was using a template marked "2022" when the year is 2023. 
-The way you type in the renewal policy term must be "Month #, 202# - Month #, 202#".
-When dragging in the policy package, it is important that you release the cursor anywhere 
in the email body (not the header image). The Ctrl F/H logic used does not work if the cursor is on the image
and the replacement of the due date only works if the cursor is in the body when it starts.
-Before unpausing on the rate sheet for storing building and contents, make sure the cursor is near the Building cell
so that it selects the first instance of the building and contents cells. It will technically work if the cursor is at the bottom, however
certain types of rate sheets have multiple instances of "Building" and "Contents" cells so it may not pick the one you want.
-If this script is being used by anyone else or the file path of your "Renewal Policies" folder changes,
the file paths need to be updated for that specific user
-Pay attention to the day number of the renewal date. If it is 31st, the script will make the three-pay plan
all 31st of those months, and some months like September might not have a 31st. Would have to fix that before
unpausing at the point where the renewal letter is print-to-pdfed
-When creating the renewal letter, the year on the attachment will enter the end number. Ex if the template is named
"SP - Renewal Letter 2022", the program will name it "SP - Renewal Letter 2023" because the last number of the current 
year is 3 (for 2023). If this is still being used way in the future when we are in 2030s, it should still get the correct
year, but uncertain at the current time. 
*/

;---------------Hotkey for AB CHUR Clients---------------
^+0::
setTitleMatchMode RegEx
;When entering all of the inputs, open the policy document to verify and enter the policy term

;---------------User input for workflow elements---------------
InputBox, client_code, Client Code, Enter client code
InputBox, policy_number, Policy Number, Enter the client's policy number
InputBox, policy_term, Policy Term, Enter the renewal policy term for the client
InputBox, cert_copy, Certified Copy, Is there a cert copy? Enter 'y' or 'n'

;---------------Filling out subject line of template & attach policy---------------
WinActivate Account Locate
Send, %client_code%{Enter}!ap
WinWaitActive ^%client_code%
Sleep 100
WinGetActiveTitle, client_window
client_name := SubStr(client_window, 14)
WinActivate .*%A_YYYY% Insurance Renewal
Send, {Tab 2}{Control Down}{Shift Down}{Right 2}{Shift Up}{Control Up}{Backspace}
Send, {Down}{Left 2}%client_code% - %policy_number%{Down}%client_name%
Pause
;it is at this point that you drag over the attachments from Tricia's email

;---------------A&S Certificate if the client has one---------------
WinActivate ^%client_code%
InputBox, A_S, Accident & Sickness Policy, Does the client have Accident and Sickness? Enter 'y' or 'n'
if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
	InputBox, A_SPremium, Accident and Sickness Premium, Enter the Accident and Sickness premium before tax (Ex: 198.00)
	InputBox, premium, Account Balance, Enter the amount due for the renewal (Ex: 1460.55 or 5743.00)
	Send, !aa
	Sleep 500
	Send, {Tab 2}d{Tab 2}SP - Acc{Enter 2}
	Pause
	Send, !fpp
	Pause
	Send, Accident Certificate
	Sleep 100
	Send, !sy
	Sleep 500
	Send, !{F4}{Enter}
	WinActivate %A_YYYY% Insurance Renewal
	Send, !nafb
	Send, Accident Certificate
	Pause
	Send, {Enter}
	Sleep 500
	WinActivate ^%client_code%
	Send, {Tab 5}{Enter}
	Sleep 500
}
else {
	InputBox, premium, Total Premium Owing, Enter the amount due for the renewal (Ex: 1460.55 or 5743.00)
	Send, !aa
	Pause
}

;---------------Saving and attaching invoice to email---------------
Send, !c{Tab}d{Tab 2}Invoice{Enter 2}
Pause 
/*
Verify that the most recent invoice is opened here and also confirm the amount in the 
invoice matches the red amount in Epic
*/
SendMode Event
Send, {LAlt Down}fd{LAlt Up}
SendMode Input
Send, {Tab 2}{Enter}
Pause
Send, Invoice
Sleep 50
Send, {F4}^a
Sleep 100
Send, C:\Users\LandonSmyth\Downloads\Renewal Policies
Sleep 100
Send, {Enter}
Pause
Send, !sy
Sleep 500
Send, !{F4}
WinActivate %A_YYYY% Insurance Renewal
Send, !nafb
Sleep 500
Send, Invoice
Pause 
Sleep 100
Send, {Enter}
Sleep 500

;---------------Edit template based on client's coverage and renewal date---------------
cut_position := RegExMatch(policy_term, "-") 
renewal_date := SubStr(policy_term, 1, cut_position-2) 
month_array := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
month_to_number := {"January":1, "February":2, "March":3, "April":4, "May":5, "June":6, "July":7, "August":8, "September":9, "October":10, "November":11, "December":12}
month_cut := RegExMatch(renewal_date, "( \d,| \d\d,)")
month := SubStr(renewal_date, 1, month_cut-1)
month_number := month_to_number[month]
Send, ^h%month% , %A_YYYY%{Tab}%renewal_date%
Sleep 100
Send, !a
Sleep 100
Send, {Enter}{Escape}
Sleep 300
WinActivate ^%client_code%
Send, {Tab 5}{Enter}
Sleep 500
Send, !c{Tab}f{Tab 2}spr{Enter 2}
;Pause
WinWaitActive Excel$
Sleep 1500 ;this likely would need tuning depending on the system or the loading time of Excel
Send, !hzefdg
Sleep 500
Send, b33{Enter} ;this should put you right above the building and contents bars to ensure it enters in the right spots
Pause ;this is necessary in case the rate sheet is something such as multi construction or multi location
A_Clipboard := ""
Send, ^fBuilding!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	building := false
} else {
	building := true
}
A_Clipboard := ""
Send, ^fContents!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	contents := false
} else {
	contents := true
}
Send, !{F4}{Right}{Enter}
Sleep 200
WinActivate %A_YYYY% Insurance Renewal
if (building and not contents){
	Send, ^hyour Building & Contents limits{Tab}your Building limit!a
	Send, {Enter}
	Sleep 300
	Send, {Escape}
}
else if (not building and contents){
	Send, ^hyour Building & Contents limits{Tab}your Contents limit!a
	Send, {Enter}
	Sleep 300
	Send, {Escape}
}
else if (not building and not contents){
	Send, ^fYour renewal has been!f{Escape}{Control Down}{Shift Down}{Down 2}{Shift Up}{Control Up}{Backspace}
}
if (cert_copy = "n" or cert_copy = "n " or cert_copy = " n" or cert_copy = "N" or cert_copy = "N " or cert_copy = " N"){
	Send, ^fA certified copy!f{Escape}+{Down}{Backspace 2}
}
Sleep 200

;---------------Create renewal letter---------------
WinActivate ^%client_code%
Send, {Tab 5}{Enter}
Sleep 500
Send, ^n
Sleep 500 ;this might be variable loading time, might have to change to pause
Send, {Enter}
Sleep 500
Send, {Tab}SP - AB Ren{Tab 2}{Space}{Down}{Space}{Tab 5}{Right 4}{Tab 2}{Space}
Pause ;here you verify that everything was selected correctly
Send, {Enter}
Sleep 400 
year = %A_YYYY%
year_number := SubStr(year, 0)
Send, {Right}{Backspace}%year_number%
Pause ;change to a sleep later once you trust it
Send, {Enter}
Pause ;here you wait for it to open, and you must hit the enable content on the macros
Send, %policy_term%{Tab}{Enter}
if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
	Send, y{Tab}{Enter}
	Sleep 100
	Send, %A_SPremium%{Tab}{Enter}
	Sleep 100
}
else {
	Send, n{Tab}{Enter}
	Sleep 100
}
Send, %renewal_date%{Tab}{Enter}
Sleep 100
Send, %premium%{Tab}{Enter}
Sleep 100
Send, %renewal_date%{Tab}{Enter}
Sleep 100
if (month_number < 11){
	month_two := StrReplace(renewal_date, month, month_array[month_number+1])
	month_three := StrReplace(renewal_date, month, month_array[month_number+2])
}
else if (month_number = 11){
	month_two := StrReplace(renewal_date, month, month_array[12])
	month_three := StrReplace(renewal_date, month, month_array[1])
}
else if (month_number = 12){
	month_two := StrReplace(renewal_date, month, month_array[1])
	month_three := StrReplace(renewal_date, month, month_array[2])
}
Send, %month_two%{Tab}{Enter}
Sleep 100
Send, %month_three%{Tab}{Enter}
Sleep 100
Send, y{Tab}{Enter}
Sleep 300
if (premium >= 30000.00){
	Send, ^h$%premium%!f{Escape}{Left}{Right 3},
	Send, {Down 2},{Down},{Down},
	if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
		Send, {PgDn}{Backspace}{Down}{Backspace}{PgUp}
	}
	
}
else if (30000.00 >= premium and premium >= 10000.00){
	Send, ^h$%premium%!f{Escape}{Left}{Right 3},
	Send, {Down 2}{Left},{Down},{Down},
	if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
		Send, {PgDn}{Backspace}{Down}{Backspace}{PgUp}
	}
}
else if (10000.00 >= premium and premium >= 1000.00) {
	Send, ^h$%premium%!f{Escape}{Left}{Right 2},
	if (premium >= 3000.00) {
		Send, {Down 2},{Down},{Down},
		if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
			Send, {PgDn}{Backspace}{Down}{Backspace}{PgUp}
		}
	}
	else {
		if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
			Send, {PgDn}{Backspace}{Down}{Backspace}{PgUp}
		}
	}
}
else {
	if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
		Send, {PgDn 2}{Right}{Backspace}{Down}{Backspace}{PgUp}
	}
}
Pause

;---------------Attach Renewal Letter to Email---------------
Send, !fpp
Pause
Send, Renewal Letter
Sleep 100
Send, !sy
Sleep 1000
Send, !{F4} 
Pause ;wait for Save Document prompt to show up
Send, {Enter}
WinActivate %A_YYYY% Insurance Renewal
Send, !nafb
Send, Renewal Letter
Pause
Send, {Enter}	
Sleep 200
Send, !ps{Enter}!pd
Send, !bb{Tab 2}8:30 AM
Pause ;here you choose the delay-send date
Send, {Enter}
Sleep 100
Send, ^hDear ,!f{Escape}{Right}{Left}
Return ;here is where you type the name, paste the email(s), and make any last changes before sending (like a statement document if needed)


;---------------Hotkey for DB CHUR Clients---------------
^+9:: 
setTitleMatchMode RegEx
;When entering all of the inputs, open the policy document to verify and enter the policy term

;---------------User input for workflow elements---------------
InputBox, client_code, Client Code, Enter client code
InputBox, policy_number, Policy Number, Enter the client's policy number
InputBox, renewal_date, Renewal Date, Enter the client's renewal date
InputBox, cert_copy, Certified Copy, Is there a cert copy? Enter 'y' or 'n'

;---------------Filling out subject line of template & attach policy---------------
WinActivate Account Locate
Send, %client_code%{Enter}!ap
WinWaitActive ^%client_code%
Sleep 100
WinGetActiveTitle, client_window
client_name := SubStr(client_window, 14)
WinActivate .*%A_YYYY% Insurance Renewal
Send, {Tab 2}{Control Down}{Shift Down}{Right 2}{Shift Up}{Control Up}{Backspace}
Send, {Down}{Left 2}%client_code% - %policy_number%{Down}%client_name%
Pause
;it is at this point that you drag over the attachments from Tricia's email

;---------------A&S Certificate if the client has one---------------
WinActivate ^%client_code%
InputBox, A_S, Accident & Sickness Policy, Does the client have Accident and Sickness? Enter 'y' or 'n'
if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
	Send, !aa
	Sleep 500
	Send, {Tab 2}d{Tab 2}SP - Acc{Enter 2}
	Pause
	Send, !fpp
	Pause
	Send, Accident Certificate
	Send, !sy
	Sleep 500
	Send, !{F4}{Enter}
	WinActivate %A_YYYY% Insurance Renewal
	Send, !nafb
	Send, Accident Certificate
	Pause
	Send, {Enter}
	Sleep 500
	WinActivate ^%client_code%
	Send, {Tab 5}{Enter}
	Sleep 500

	;---------------Saving and attaching A&S invoice to email---------------

	Send, !c{Tab}d{Tab 2}Invoice{Enter 2}
	Pause 
	/*
	Verify that the most recent invoice is opened here and also confirm the amount in the 
	invoice matches the red amount in Epic
	*/
	SendMode Event
	Send, {LAlt Down}fd{LAlt Up}
	SendMode Input
	Send, {Tab 2}{Enter}
	Sleep 500
	Send, Invoice
	Sleep 50
	Send, {F4}^a
	Sleep 100
	Send, C:\Users\LandonSmyth\Downloads\Renewal Policies
	Sleep 100
	Send, {Enter}
	Sleep 1000
	Send, !sy
	Sleep 500
	Send, !{F4}
	WinActivate %A_YYYY% Insurance Renewal
	Send, !nafb
	Sleep 500
	Send, Invoice
	Pause
	Send, {Enter}
	Sleep 700
}
else {
	Sleep 300
	Send, !aa
	Sleep 300
	WinActivate %A_YYYY% Insurance Renewal
	Send, ^h*An invoice is also!f{Escape}{Control Down}{Shift Down}{Down 2}{Shift Up}{Control Up}{Backspace}
	Sleep 100
}

;---------------Edit template based on client's coverage and renewal date---------------
month_cut := RegExMatch(renewal_date, "( \d,| \d\d,)")
month := SubStr(renewal_date, 1, month_cut-1)
Send, ^h%month% , %A_YYYY%{Tab}%renewal_date%!a{Enter}{Escape}
Sleep 300
WinActivate ^%client_code%
if (A_S = "y" or A_S = "y " or A_S = " y" or A_S = "Y" or A_S = "Y " or A_S = " Y"){
	Send, {Tab 5}{Enter}
	Sleep 500
}
Send, !c{Tab}f{Tab 2}spr{Enter 2}
;Pause
WinWaitActive Excel$
Sleep 1800 ;this likely would need tuning depending on the system or the loading time of Excel
Send, !hzefdg
Sleep 500
Send, b33{Enter} ;this should put you right above the building and contents bars to ensure it enters in the right spots
Pause ;this is necessary in case the rate sheet is something such as multi construction or multi location
A_Clipboard := ""
Send, ^fBuilding!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	building := false
} else {
	building := true
}
A_Clipboard := ""
Send, ^fContents!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	contents := false
} else {
	contents := true
}
Send, !{F4}{Right}{Enter}
Sleep 200
WinActivate %A_YYYY% Insurance Renewal
Sleep 50
Send, {PgUp}
if (building and not contents){
	Send, ^hyour Building & Contents limits{Tab}your Building limit!a
	Send, {Enter}
	Sleep 300
	Send, {Escape}
}
else if (not building and contents){
	Send, ^hyour Building & Contents limits{Tab}your Contents limit!a
	Send, {Enter}
	Sleep 300
	Send, {Escape}
}
else if (not building and not contents){
	Send, ^fYour renewal has been!f{Escape}{Control Down}{Shift Down}{Down 2}{Shift Up}{Control Up}{Backspace}
}
if (cert_copy = "n" or cert_copy = "n " or cert_copy = " n" or cert_copy = "N" or cert_copy = "N " or cert_copy = " N"){
	Send, ^fA certified copy!f{Escape}+{Down}{Backspace 2}
}
Sleep 200
Send, !ps{Enter}!pd
Send, !bb{Tab 2}8:30 AM
Pause ;here you choose the delay-send date
Send, {Enter}
Sleep 100
Send, ^hDear ,!f{Escape}{Right}{Left}
Return ;here is where you type the name, paste the email(s), and make any last changes before sending (like a statement document if needed)



;Testing key

^+1::
setTitleMatchMode RegEx
Send, !c{Tab}f{Tab 2}spr{Enter 2}
;Pause
WinWaitActive Excel$
Sleep 1700
Send, !hzefdg
Sleep 500
Send, b33{Enter} ;this should put you right above the building and contents bars to ensure it enters in the right spots
Pause ;this is necessary in case the rate sheet is something such as multi construction or multi location
A_Clipboard := ""
Send, ^fBuilding!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	building := false
} else {
	building := true
}
A_Clipboard := ""
Send, ^fContents!f
Sleep 200
Send, {Escape}{Right}{F2}{Shift Down}{Up}{Shift Up}^c
Sleep 300
if (A_Clipboard = ""){
	contents := false
} else {
	contents := true
}
Return

^+2::
cert_copy := "n"
if (cert_copy = "n" or cert_copy = "n " or cert_copy = " n" or cert_copy = "N" or cert_copy = "N " or cert_copy = " N"){
	Send, ^fA certified copy!f{Escape}+{Down}{Backspace 2}
}
Return