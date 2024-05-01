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
-The program deletes the first attachment which is usually the policy document. If it is anything different, 
it will remove it all the same.
-Replace logic will likely choose the wrong etransfer email if there are more than one reminders in the thread. 
Will need tweaking if using for 2nd 3rd and higher number reminders.
*/
;---------------Before Running: Open the client by clicking on the activity in Epic Home---------------

;---------------1st Reminder---------------
^+1::
template := "
(
I’m emailing to follow up with you on payment for client name’s Insurance Renewal (email below). We have not received it yet. Payment was due in our office by payment date. Please advise once payment has been sent or if it has already been sent. I have attached a copy of the Invoice again for your reference.

We also have an Interac E-Transfer Payment Option available through online banking.

If using the e-transfer option to make the payment of $amount please follow the steps below:
	1. Go into your online banking website
2. Choose `""Interac e-Transfer`""
3. Email to: etransfer@brokerforce.ca, please identify your lookup code (client code) in the comments section


Thank you for your prompt attention to this matter.
)"
setTitleMatchMode RegEx
WinGetActiveTitle, client_window
client_code := SubStr(client_window, 1, 7)
client_name := SubStr(client_window, 14)
Send, !aa
Pause
Send, {Tab 2}d{Tab 2}Insurance Renewal{Enter 2}
WinWaitActive ^\d\d\d\d Insurance Renewal
Sleep 100
Send, ^a{Tab 5}
Sleep 100
Send, +{Up}^c
InputBox, recipient_names, Names of Recipients, Enter the names of the email recipients (Ex: Mike and Paul)
Send, ^f
Sleep 500 ;decrease if possible
Send, ^v{Tab 2}{Control Down}{Shift Down}{Right}{Shift Up}{Control Up}
Send, Payment Reminder{Tab}{Del}{Tab} ;
Sleep 100
right_cut := RegExMatch(client_window, " - ")
client_code := SubStr(client_window, 1, right_cut-1)
WinActivate ^%client_code%
Sleep 100
Send, !ap
InputBox, payment_date, Due Date of Payment, Enter the due date of the payment (Ex: July 14th)
InputBox, balance, Current Balance, Enter the the current balance for the payment (Ex: 1`,432.35 or 4`,032.00)
Sleep 200
Send, !aa
Sleep 300
Send, {Tab 5}{Enter}
WinActivate ^Payment Reminder
Sleep 100
Send, Dear %recipient_names%,{Enter 2}%template%
Sleep 600
Send, ^h
Sleep 100
Send, client name{Tab}%client_name%!a{Enter}
Sleep 50
Send, payment date{Tab}%payment_date%^u!a{Enter}
Sleep 50
Send, $amount{Tab}$%balance%^u^b!a{Enter}
Sleep 50
Send, client code{Tab}%client_code%^b!a{Enter}{Escape}
Sleep 100
Send, ^fetransfer@brokerforce.ca!f{Escape}^b
Send, {Right}^{Right}{Shift Down}{Down}{Left 2}{Shift Up}^u^b^!h{PgUp}
Sleep 100
Send, ^h
Sleep 100
Send, xyz{Tab}^u^b!f{Enter}{Escape}
Return


;testing
^+0::
template := "
(
I’m emailing to follow up with you on payment for client name’s Insurance Renewal (email below). We have not received it yet. Payment was due in our office by payment date. Please advise once payment has been sent or if it has already been sent. I have attached a copy of the Invoice again for your reference.

We also have an Interac E-Transfer Payment Option available through online banking.

If using the e-transfer option to make the payment of $amount please follow the steps below:
	1. Go into your online banking website
2. Choose `""Interac e-Transfer`""
3. Email to: etransfer@brokerforce.ca, please identify your lookup code (client code) in the comments section


Thank you for your prompt attention to this matter.
)"

setTitleMatchMode RegEx
client_name := "New Life in Christ Church"
client_code := "NEWLI01"
recipient_names := "Rob"
payment_date := "July 1st"
balance := 1345.00
Send, Dear %recipient_names%,{Enter 2}%template%
Sleep 600
Send, ^h
Sleep 100
Send, client name{Tab}%client_name%!a{Enter}
Sleep 50
Send, payment date{Tab}%payment_date%^u!a{Enter}
Sleep 50
Send, $amount{Tab}$%balance%^u^b!a{Enter}
Sleep 50
Send, client code{Tab}%client_code%^b!a{Enter}{Escape}
Sleep 100
Send, ^fetransfer@brokerforce.ca!f{Escape}^b
Send, {Right}^{Right}{Shift Down}{Down}{Left 2}{Shift Up}^u^b^!h{PgUp}
Sleep 100
Send, ^h
Sleep 100
Send, xyz{Tab}^u^b!f{Enter}{Escape}
Return

