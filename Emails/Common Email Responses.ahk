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
*/

^+0::
InputBox, template, Email Template, Enter "abuse"/"coi"/"outside groups"/"inflation"/"received" based on the email template you need
if (template = "abuse" or template = "Abuse" or template = "abuse " or template = "Abuse "){
	Send, As part of an approved protocol, the church needs to meet all requirements outlined in the protocol (Questions 16 – 22 on the abuse section). Please note that {Enter 2}To confirm that you still qualify for abuse coverage, we will need you to confirm the items will be implemented going forward. Please send a revised application with the responses changed from “no” to “yes” and initial beside each revised answer.
}
else if (template = "coi" or template = "Coi" or template = "coi " or template = "Coi "){
	Send, Please provide the below information for the certificate at your earliest convenience:{Enter 2}
	Send, •	Name of the Facility being used:{Enter}
	Send, Address of the facility being used:{Enter}
	Send, Details of the event held at this facility:{Enter}
	Send, Date(s) you will be using this facility:{Enter 3}Once I have this information, I can issue the certificate to your attention. Please let me know if you have any questions.{Enter}
}
else if (template = "outside groups" or template = "Outside groups" or template = "outside groups " or template = "Outside groups "){
	Send, Please ensure that you are obtaining a certificate of liability insurance from each outside group with a minimum $2,000,000 liability coverage naming “ as an additional insured”.
	Send, {Control Down}{Left 5}{Control Up}{Left}
}
else if (template = "inflation" or template = "Inflation" or template = "inflation " or template = "Inflation "){
	Send, {Text}Please note that upon renewal we apply a 5`% inflationary increase to your property limits to account for the rising cost of construction and consumer goods.
}
else if (template = "received" or template = "Received" or template = "received " or template = "Received "){
	Send, Thank you for sending the completed renewal questionnaire and financials, I have added the documents to your file. 
	Send, {Down 2}{Left 2}{Control down}{Shift down}{Left 2}{Shift up}{Control up}Have a great day
}
Return







