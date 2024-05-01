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
-The logic for opening the rate sheet from the SP Shared folder requires that the month folder be open 
and the cursor be on one of the files in the folder. If it is not, it will not snap to the correct rate sheet
-If Epic changes the arrangement of where the premium input boxes are under the policy tab, the tab count(s) will needed updated
-After you enter the payment type as "ab" or "db", the program will right away select the policy line one beneath the top
policy. This is 9 times out of 10 the CHUR line that has no premium (Good). If the client has event or auto insurance at the 
top, *before* entering the payment type you must click on the policy above the CHUR line for the current term.
-This program assumes that you will never have a policy that does not have a contents limit (liability only client). If you get 
one of these, make sure to stop the script when you see it and send to Nithila. 
-You cannot have decimals in numbers that you enter into PW, it considers the decimal part as two extra figures
-Currently the script only updates an EBI limit in PW if the client has a building and contents limit.
If they have only contents limit(s) but have EBI in PW, it has to be manually entered before continuing 
-The sequence for attaching the rate sheet assumes that the top policy is the most recent CHUR policy.
If they have auto or event insurance, it would need to be changed before unpausing for entering it
-Navigation in Foxit to the summary of changes and deductibles pages relies upon {Right} incrementing the page number 
by one. If Foxit ever changes this keybind, it will need updated to either many Page Down key presses, or the 
new corresponding keybind
-The porgram assumes that a client that has a building and contents limit also has an EBI limit to update.
In cases where the client has both property types but no EBI, the program will send the EBI limit as keystrokes and 
it will need removed. Could be fixed by adding another input box to True of False EBI, but not worth it
until determined the frequency of this type of case. 
-If you reload before using the key to attach the rate sheet and send email to Tricia, you must do it manually.
Many parts of that key rely on previously defined varaibles from the first key and if you reload, it will 
not sucessfully window match and will have parts of the email missing

If you go to print and there is an alarm task
go to risks -> fire -> crime -> choose monitoring service (shared service)
The one that needs updated will be marked Central Station
Always do Apply then Ok Then Save
*/

;---------------Hotkey for Updating Limits and confirming details from NB correct---------------
^+0::
setTitleMatchMode RegEx
SetKeyDelay, 50 ;this must be here to allow policy works to keep up with input
InputBox, client_code, Client Code, Enter client code
Send, %client_code%{Enter}!av
Pause ;here you confirm that there are no important activities open that would stop you from updating
WinGetActiveTitle, client_window
if(InStr(client_window, "(") != 0 or InStr(client_window, ")") != 0){
	client_window := SubStr(client_window, 1, 13)
}
client_name_string := SubStr(client_window, 14, 5)
Send, !ap
Sleep 200
;should select one above the CHUR line that has the premium, if not, select it BEFORE entering payment type
InputBox, payment_type, Payment Type, Enter the client's payment type (Ex: "ab" or "db")
Send, {Down}+{F2}
Sleep 200
Send, {Tab 2}{Down 5}{Enter}
WinWaitActive ^RemoteApp
Sleep 500
WinActivate ^%client_code%
Send, !aa
Pause ;here is where you open the policy documents to get what you need 
cut_position := RegExMatch(client_window, " - ") 
month := SubStr(client_window, 8, cut_position-8) 
WinActivate - %month%
Sleep 100
SendMode Event
Send, %client_code%
SendMode Input
Sleep 200
Send, {Enter}
Sleep 1600
WinActivate (?<!Comment )\(Applied Cloud\)$
Pause ;here is where you open the comision doc and have PW open on a main screen
InputBox, previous_premium, Previous year premium, Enter the previous year premium shown in Policy Works
InputBox, renewal_premium, Renewal premium, Enter the renewal premium shown in the NB comission document
increase_cap := Round(1.06 * previous_premium)
if (renewal_premium <= increase_cap){
	MsgBox, 262144, Check Within 6`% Increase,Previous Premium: %previous_premium%`nRenewal Premium: %renewal_premium%`n1.06 x %previous_premium% = %increase_cap%`n`nRenewal premium is within 6`% increase, continue
	WinActivate (?<!Comment )\(Applied Cloud\)$
	Sleep 500 ;this might need increased or changed into a pause
	Send, ^f
	Sleep 500
	Send, Sanctuary Plus Religious!f
	Sleep 600
	Send, {Escape}
	Sleep 500
	Send, {Right 6}
	Pause ;cursor should be on the premium in the top right
	SendMode Event
	Send, %renewal_premium%{Enter}
	SendMode Input
	Sleep 500
	WinActivate %client_code% R \d\d
	Send, ^fRevenue!f
	Sleep 200
	Send, {Escape}{Right}{F2}+{Up}^c
	Sleep 200
	revenue := A_Clipboard
	Send, {Escape}
	Sleep 100
	Send, ^f2023 Renewal Premium!f
	Sleep 200
	Send, {Escape}{F2}$%renewal_premium% ;NOTE need to remove the $ once you are in a month that was not done by Nithila
	if (renewal_premium < 1000){
		Send, {Enter}
	}
	else {
		Send, {Left 3},{Enter}
	}
	Pause ;change to sleep 500 when functional
	Send, {PgUp}
	Sleep 100
	WinActivate %client_window%
	Sleep 100
	Send, !ap
	Pause ;cursor should be on the CHUR policy that has the premium, beneath the one to be finalized
	Send, {Up}
	Sleep 500
	Send, {Enter}
	Pause
	Send, !c{Tab}%renewal_premium%{Tab}
	Pause ;here is where you click on the "Line" tab on the left side of the window
	if (payment_type = "ab" or payment_type = "AB"){
		Send, {Tab 19}
	}
	else {
		Send, {Tab 18}
	}
	Sleep 100
	Send, %renewal_premium%
	Sleep 500
	Send, {Tab 5}{Enter}
	Pause ;must manually click the 'X' on the left hand side to finalize the line
	Send, {Down}
	Sleep 200
	Send, {Tab}
	Sleep 100
	Send, Iss
	Sleep 200
	Send, {Tab}
	Pause
	Send, {Enter}
	WinWaitActive ^Issue
	Sleep 200
	Send, {Enter}
	WinActivate Foxit PDF Reader$
	Send, !{F4}
	Sleep 700 ;tune if needed
	Send, u
	Sleep 1300
	SetKeyDelay, 75
	SendMode Event
	Send, {Right 3}{PgDn} ;this should bring you to the summary of changes page in the policy document
	SendMode Input
	SetKeyDelay, 50
	WinActivate %client_code% R \d\d
	InputBox, building_limit, Enter New Building Limit, Enter the renewal-term building limit (Enter '0' if they do not have a building)
	InputBox, contents_limit, Enter New Contents Limit, Enter the renewal-term contents limit
	Sleep 100
	WinActivate %client_code% R \d\d
	Sleep 100
	Send, !hzefdg
	Sleep 500
	Send, b33{Enter} ;this should put you right above the building and contents bars to ensure it enters in the right spots
	Sleep 100
	if (building_limit != "0"){
		Send, ^fBuilding!f
		Sleep 200
		Send, {Escape}{Right}%building_limit%{Enter}
		Sleep 200 ;assumes that pressing enter puts the cursor on the contents limit box
		Send, %contents_limit%{Enter}
		Sleep 200
		EBI := building_limit + contents_limit
		MsgBox, 0, EBI Limit, EBI = %building_limit% + %contents_limit% = %EBI%`nConfirm rate sheet matches
		;here is when you would make any additional updates like contents limits for multiple locations
		Pause ;confirm here that the EBI limit in the rate sheet is the sum of the building and contents limits
		WinActivate (?<!Comment )\(Applied Cloud\)$
		Sleep 500
		Send, ^f
		Sleep 500
		Send, Building!f
		Sleep 600
		Send, {Escape}
		Sleep 500
		Send, {Right 4}
		Pause ;cursor should be on the previous term building limit 
		SendMode Event
		Send, %building_limit%{Enter}
		SendMode Input
		Sleep 500

		;NOTE: you might be able to have it enter building and then just {Down} to select the contents. Unsure if this is always where contents will be
	
		Send, ^f
		Sleep 500
		Send, Business Personal Property!f
		Sleep 600
		Send, {Enter}
		Sleep 500
		Send, {Escape}
		Sleep 500
		Send, {Right 4}
		Pause ;cursor should be on the previous term contents limit 
		SendMode Event
		Send, %contents_limit%{Enter}
		SendMode Input
		Sleep 500

		Send, ^f
		Sleep 500
		Send, Equipment Breakdown End!f
		Sleep 600
		Send, {Enter}
		Sleep 500
		Send, {Escape}
		Sleep 500
		Send, {Right 4}{Down}
		Pause ;cursor should be on the EBI limit 
		SendMode Event
		Send, %EBI%{Enter}
		SendMode Input
		Sleep 500
		Pause ;update any additional location contents/building limits
	}
	else {
		Send, ^fContents!f
		Sleep 200
		Send, {Escape}{Right}%contents_limit%{Enter}
		Sleep 200
		Pause ;here is when you would make any additional updates like contents limits for multiple locations or adjust EBI
		WinActivate (?<!Comment )\(Applied Cloud\)$
		Sleep 500
		Send, ^f
		Sleep 500
		Send, Business Personal Property!f
		Sleep 600
		Send, {Escape}
		Sleep 500
		Send, {Right 4}
		Pause ;cursor should be on the previous term contents limit 
		SendMode Event
		Send, %contents_limit%{Enter}
		SendMode Input
		Sleep 500
		Pause ;update any additional location contents/building limits
	}
	WinActivate Foxit PDF Reader$
	Sleep 300
	SendMode Event
	Send, {Right 2}
	SendMode Input
	WinActivate (?<!Comment )\(Applied Cloud\)$
	
	Pause ;check all deductibles 
}
else {
	MsgBox, 0, Check Within 6`% Increase,Previous Premium: %previous_premium%`nRenewal Premium: %renewal_premium%`n1.06 x %previous_premium% = %increase_cap%`n`nRenewal premium above 6`% increase, send to Nithila
	Return
}
;before unpausing you must check all of the limits/deductibles 
WinActivate Foxit PDF Reader$
Sleep 100
Send, !{F4}
Sleep 200
WinActivate (?<!Comment )\(Applied Cloud\)$
/*
Sleep 500
Send, ^f
Sleep 500
Send, Forest Fire Fighting!f
Sleep 600
Send, {Escape}
Sleep 500
Send, {Right 4}
Sleep 100
SendMode Event
Send, 500000{Enter}
SendMode Input
*/
Sleep 500
Send, ^s
Pause ;once all of the limits have been updated, continue
Send, !uu
Pause ;must click on liability and highlight revenue input field
Send, %revenue%
Pause ;here is where you press Apply in the bottom right
Send, !up
Sleep 500
Send, {Enter}
;here you click the golden triangle
Return

;---------------Hotkey for calculating 5% increase values for building and contents limits---------------
^+5::
InputBox, previous_building, Previous Building Limit, Enter the previous year building limit. Enter '0' if they have no building
InputBox, previous_contents, Previous Contents Limit, Enter the previous year contents limit
renewal_building := Round(1.05 * previous_building, 2)
renewal_contents := Round(1.05 * previous_contents, 2)
EBI := Round(renewal_building + renewal_contents, 2)

;adding commas to the input strings
CommaSeparate(number, separator=",") {
	return RegExReplace(number, "\G\d+?(?=(\d{3})+(?:\D|$))", "$0" separator) ;got this from a forum online 
}
previous_building := CommaSeparate(previous_building)
previous_contents := CommaSeparate(previous_contents)
renewal_building := CommaSeparate(renewal_building)
renewal_contents := CommaSeparate(renewal_contents)
EBI := CommaSeparate(EBI)

MsgBox, 262144, Inflation Increase Calculator,Previous Building: %previous_building%`nIncreased building: %renewal_building%`n`nPrevious Contents: %previous_contents%`nIncreased contents: %renewal_contents%`n`nSum: %EBI%
Return

;---------------Hotkey for attaching the rate sheet an sending the renewal email to Tricia---------------
^+2::

;When running this key either make sure you do it right away after you start the policy printing 
;or do it after the policy printing is done and PW is closed. The popup asking to add an activity would 
;interfere with the keystrokes and steps of the key


setTitleMatchMode RegEx
WinActivate R \d\d
Sleep 100
Send, !{F4}
Sleep 500
WinActivate %client_window%
Sleep 100
Send, !aa
Sleep 100
WinActivate - %month%
Pause ;here you drag over the rate sheet into the attachments
Send, {F4}
Sleep 500 ;here you just double check that it is picking the top CHUR policy and not auto or event policy
Send, {Tab}{Enter}
Sleep 100
Send, {Tab 2}{Right}^{Backspace 2}{Tab 2}spre{Tab}
Pause
Send, {Enter}
while (True){
	if (FileExist("C:\Policy Works Files\*.pdf")){
		break
	}
	else {
		Sleep 300
	}
}
WinActivate (?<!Comment )\(Applied Cloud\)$
Sleep 1000
Send, !{F4}
Sleep 100
WinActivate %client_window%
Send, !c{Tab 7}{Enter}
Pause ;here you just eyeball that it actually made the link to PW correct in their file
Send, !{F4}
WinActivate Outlook$
Sleep 100
Send, !hn1
Sleep 500 ;might need adjusted depending on the speed of the device
Send, t.fraser@brokerforce.ca
Pause ;tabbing out the email is slow and never worked smoothly without the pause so rather keep it in for reliability
Send, {Enter}
Sleep 50
Send, {Tab 2}%client_code%%month% Renewal{Tab}
Sleep 100
Send, !nafb
Sleep 600 ;might need tuning
Send, {F4}^a
Sleep 100
Send, C:\Policy Works Files
Sleep 100
Send, {Enter}
Sleep 500 ;this may need increased
Send, !n%client_name_string%{Down}{Enter}
Sleep 1500
FileCopy, C:\Policy Works Files\*.pdf, C:\Policy Works Files (Backup of print history)
FileDelete, C:\Policy Works Files\*.pdf
Return


;---------------Testing---------------
^+1::
setTitleMatchMode RegEx
while (True){
	if (FileExist("C:\Policy Works Files\*.pdf")){
		MsgBox There is a pdf in the folder, break
		break
	}
	else {
		Sleep 500
	}
}
FileCopy, C:\Policy Works Files\*.pdf, C:\Policy Works Files (Backup of print history)
FileDelete, C:\Policy Works Files\*.pdf
Return

^+y::
SetKeyDelay, 75
WinActivate Foxit PDF Reader$
Send, !{F4}
Sleep 700 ;tune if needed
Send, u
Sleep 1300
SendMode Event
Send, {Right 3}{PgDn} ;this should bring you to the summary of changes page in the policy document
SendMode Input
MsgBox Done Bozo
Return