#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

`::
Pause -1
Return

PgUp::
reload
Return

activate_ahk_windows(){
	setTitleMatchMode RegEx
	while (True){
		IfWinExist Renewal premium
		{
			Sleep 200
			WinActivate Renewal premium
			Sleep 2000
		}
		IfWinExist ^Check Within 6
		{
			Sleep 200
			WinActivate ^Check Within 6
			Sleep 2000
		}
		else {
			Sleep 500
		}
	}	
}

^+3::
activate_ahk_windows()
Return


^+1:: 
Send, !aa
Sleep 3000
Send, {Tab 2}d{Tab 2}SP - Acc{Enter}
Sleep 500
Send, {Enter}
Sleep 4000
Send, ^h2024-{Tab}2024-{Enter}!r{Enter}{Right}{Left}{Backspace}2
, {Tab}{Right}{Left}{Backspace}3{Enter}- 2{Tab}- 2{Enter}!r{Enter}{Escape}

^+2::

;--------------------Open the email template--------------------

Send, This is the test
Sleep 3000

;--------------------Wait 5 seconds--------------------

Send, three seconds has passed
Sleep 2000

;--------------------alshasdfosudfh--------------------
Send, now 5 seconds has passed
Sleep 1000




Send, reached the end with a beeg whitespace
Return

^+0::
setTitleMatchMode RegEx
WinActivate ^Book1
Send, ^fD&O!f{Escape}{Right} ^z^c!{F4}
;if copying gets just the text and not an image:
;Send, ^fD&O!f{Escape}{Right}^c!{F4}
WinActivate Discord$
if (A_Clipboard = "2000000" or A_Clipboard = "5,000,000" or A_Clipboard = "10,000,000"){
	Send, success, they have D&O ;add logic to delete D&O paragraph 
}
else {
    Send, no D&O 
}
A_Clipboard := ""


^+d::
Abuse_D_O := "
(
Good 

This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire. Send it in along with copies of the latest financial statements and list of Directors and Officers at the Church. If the financials are not yet available, please send the questionnaire separately for the time being and let me know when the financials are expected to be ready/finalized. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Additionally, as you have abuse liability coverage you are required to complete the abuse application section this year (Questions 16-23 on the questionnaire).

The insurance company is doing a review and needs to make sure that their clients with abuse coverage are following all requirements to continue qualifying for an increased abuse limit.

Questions 16 – 22 on the application are the requirements that need to be included in your protocol and followed. To confirm that they are being followed, we will need you to return the form with the responses answered “yes” (excluding question 17 items 2&3).

Please let me know if you have any questions. 
)"

Abuse_Only := "
(
Good 

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
Good 

This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire. Send it in along with copies of the latest financial statements and list of Directors and Officers at the Church. If the financials are not yet available, please send the questionnaire separately for the time being and let me know when the financials are expected to be ready/finalized. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Please let me know if you have any questions. 
)"

Neither_coverage := "
(
Good 

This is a friendly reminder that we have not yet received the completed renewal questionnaire.

Please complete the attached questionnaire and return it to me. 
 
The renewal questionnaire is a requirement from the Insurance Co., and will be requested every year in order to assess changes in the insurance needs of the church. This keeps us aware of the needs of churches for their coverage and limits. Our office does get audited, and the questionnaire needs to be present in your file. 

Please let me know if you have any questions. 
)"
Abuse := False
D_O := True
if (Abuse and D_O){
	Send, %Abuse_D_O%
}
else if (Abuse and not D_O){
	Send, %Abuse_Only%
}
else if (not Abuse and D_O){
	Send, %D_O_Only%
}
else if (not Abuse and not D_O){
	Send, %Neither_Coverage%
}
Return


^+4::
setTitleMatchMode RegEx
WinGetActiveTitle, email
MsgBox %email%
while WinExist(email){
	Sleep 200
	MsgBox Window Exists still
}
MsgBox Window is Gone
Return