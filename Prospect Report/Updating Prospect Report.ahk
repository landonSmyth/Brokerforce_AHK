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
-

---------------Common Anomalies & Cases---------------
-If you find a client in CRA with matching name/City but the address is different, update the client and address. 
if the city is different, highlight on the sheet Purple
-If you find a client with a revoked charitable status, only deactivate if it is a complete match. 
If there is something different like name or address, mark Blue in the sheet 
-If you find a client in CRA that has a slightly different name than what is in Epic, update it. If the name
is significantly different, do not consider them the same and mark as yellow for "not found in CRA". 
Example: "Church of Christ" in Epic and "Church of Christ, Red Deer" are the same so update it 
-Address directions such as S, NE, SW, should not be typed out, leave them as is 
-If the address of the client has a unit number, update Epic to match whatever is listed in CRA
-If a church is incorporated and it is in their legal name, follow the naming convention in CRA. 
-For address that have RR in them, the naming convention is "RR #1", meaning the address would be written like
"123 Test Street, RR #1"
-If there is Highway in the address, the naming convention is "Highway #7". 
-If you see "General Delivery", do not omit it from the address. Add it in where it shows in CRA or in Epic originally
-If you see something like "U#C44", make it "Unit C44"
-If there is a PO box address that has a street originally in Epic but only the PO Box in CRA, omit the street address
-If you are updating the contacts and you cannot delete the old contact because it is the main business contact,
right click the new contact, make it the primary business, and now delete the old one
-if you cannot find a client that has "Bible Baptist Church" anywhere in the name, try googling to find it 
before marking it yellow -> they are likely independent and wont show in CRA 
-If you cannot find the client in CRA and in searching for it in google you see "Permanently Closed" on the right
beside the top result, inactivate the client. Andrea advised June 27, 2023 that in order for that banner
to be shown, the client must go to their google account and manually enter that they are permanently closing. 
-If you cannot find a client in CRA or google and you search by their address in google maps, find a church
at that address (that is not their name), search that name up in CRA and Epic. If it shows up in CRA with a date of 
status earlier than the date that the client was entered into Epic (such as 1967 date in CRA and 2001 in Epic account),
double check that the new name church is not a rental host by checking google and Epic. If not results for either, update the client
with the new legal name and make note of the CRC number in their account comments 

---------------Common Church Name Abbreviations---------------
BIC = brethren in christ
MB = mennonite brethren
FM = free methodist
COC = church of christ
UB = united brethren

---------------Hightlight Legend---------------
Yellow: could not find client in CRA and nothing came up in google
Blue: revoked charitable status but could not deactivate or address/city do not match
Purple: found in CRA with same name but different city and/or province address
Green: Any other anomaly to notify Andrea about 
Orange/Gold: Could not find in CRA, able to update results based on Google?

---------------Hotkey Quick Reference (All Ctrl + Shift with the respective number)---------------
0 -> Lookup the client and search for them in CRA
7 -> Make a new business contact for the clients that are set as "Personal"
2 -> Reset CRA charity search, close the client, and highlight the current row for decision/deletion
9 -> Send whatever is in the keyboard as a title-case string
3 -> Send client name and address into Epic because it is a perfect match
*/

MButton::Send, !{Left}

;---------------Hotkey to lookup client and search them in CRA---------------
^+0::
setTitleMatchMode RegEx
Send, {F2}{Shift Down}{Up}{Shift Up}^c
Sleep 100
client_code := A_Clipboard
Sleep 100
WinActivate Account Locate
Send, %client_code%{Enter}!ad
WinWaitActive ^%client_code%
WinGetActiveTitle, client_window
client_name := SubStr(client_window, 14)
IfWinExist, ^List of
{
	WinActivate ^List of
}
else {
	MsgBox Search Window not open, open it and unpause
	Sleep 500
	Pause
}
A_Clipboard := client_name
Sleep 200
Send, {Tab}{Enter}
Sleep 200
Send, {Tab}%client_name%{Tab 4}{Down 7}{Tab 7}{Enter}
Pause
Send, {PgDn}
Pause ;if there is a perfect match, click on it and unpause when the page loads
Send, {Tab}{Enter}
Sleep 200
Send, {Tab}{Enter}
WinActivate ^%client_code%
Return



;---------------Hotkey to make new Business contact---------------
^+7::
WinGetTitle, window_title, A
window_title_cut := RegExMatch(window_title, "- ")
client_name := SubStr(window_title, window_title_cut+2)
Send, !fno 
Sleep 50
Send, {Right}{Tab}%client_name%{Tab 3}{Space}
Sleep 100
Send, {Enter}
Sleep 500
Send, !ac
Sleep 100
Send, {Up}
Pause ;this is here just to make sure that the correct contact is removed 
Send, {Delete}
Sleep 100
Send, !y
Sleep 500
Send, !tp
Sleep 100
Send, !y
Sleep 500
Click, 216 201
Return


;---------------Hotkey to Reset the CRA lookup and other programs---------------
^+2::
setTitleMatchMode RegEx
;client_name := A_Clipboard
;StringLower, lower_client_name, client_name, T
;StringUpper, upper_client_name, client_name
;WinActivate (%lower_client_name%$|%upper_client_name%$)
WinActivate (?<!Inbox)(?<!Prospect)(?<!OneDrive) - (?!Landon)(?!Excel)(?!3CX)(?!Notepad)(?!Google)(?!Mozilla)[a-zA-Z]+
Send, !{F4}
WinActivate Mozilla Firefox$
Sleep 300
Send, ^l
Sleep 100
Send, https://apps.cra-arc.gc.ca/ebci/hacc/srch/pub/dsplyAdvncdSrch
Sleep 300
Send, {Enter}
WinActivate ^Prospect List
Send, {Escape}+{Space}
Return

;---------------Hotkey to type out the client name in title case---------------
^+9::
WinGetTitle, window_title, A
window_title_cut := RegExMatch(window_title, "- ")
client_name := SubStr(window_title, window_title_cut+2)
StringLower, client_name, client_name, T
Send, %client_name%
Return

;---------------Hotkey to replace all-caps name and address if it is a perfect match---------------
^+3::
/*
before starting the script, ensure the client's name is in your clipboard. Either it is already in your clipboard from
searching in CRA, or you need to copy the updated legal name from CRA before starting
*/

setTitleMatchMode RegEx
WinActivate (?<!Inbox)(?<!Prospect)(?<!OneDrive) - (?!Landon)(?!Excel)(?!3CX)(?!Notepad)[a-zA-Z]+
client_name := A_Clipboard ;this should always be the all caps version of the insured's name
StringLower, client_name, client_name, T
Send, %client_name%
WinActivate Mozilla Firefox$
Sleep 100
Send, ^f
Sleep 300
Send, Address{Escape}+{Down}^c
Sleep 100
address_untrimmed := A_Clipboard
left_trim := RegExMatch(address_untrimmed, "Address:")
right_trim := RegExMatch(address_untrimmed, "City:")
address := SubStr(address_untrimmed, left_trim+10, right_trim-15)

;---------------Replace any address abbreviations with typed-out version---------------
address := RegExReplace(address, "(?<!\w)AVE?($|(?!\w)(\.?,?))", "AVENUE") 
address := RegExReplace(address, "(?<!\w)RD($|(?!\w)(\.?,?))", "ROAD") 
address := RegExReplace(address, "(?<!\w)NO($|(?!\w)(\.?,?))", "#") 
address := RegExReplace(address, "(?<!\w)BLVD($|(?!\w)(\.?,?))", "BOULEVARD") 
address := RegExReplace(address, "(?<!\w)HWY($|(?!\w)(\.?,?))", "HIGHWAY") 
address := RegExReplace(address, "(?<!\w)STR?($|(?!\w)(\.?,?))", "STREET") 
address := RegExReplace(address, "(?<!\w)DR($|(?!\w)(\.?,?))", "DRIVE") 
address := RegExReplace(address, "(?<!\w)RANGE ROAD($|(?!\w)(\.?,?))", "RR") 
address := RegExReplace(address, "(?<!\w)CRES($|(?!\w)(\.?,?))", "CRESCENT") 
address := RegExReplace(address, "(?<!\w)CONC\.?($|(?!\w)(\.?,?))", "CONCESSION") 
address := RegExReplace(address, "(?<!\w)CR?T($|(?!\w)(\.?,?))", "COURT") 
address := RegExReplace(address, "(?<!\w)JCT($|(?!\w)(\.?,?))", "JUNCTION") 
address := RegExReplace(address, "(?<!\w)LN($|(?!\w)(\.?,?))", "LANE")
address := RegExReplace(address, "(?<!\w)PKWY($|(?!\w)(\.?,?))", "PARKWAY")
address := RegExReplace(address, "(?<!\w)PKY($|(?!\w)(\.?,?))", "PARKWAY")
address := RegExReplace(address, "(?<!\w)PL($|(?!\w)(\.?,?))", "PLACE")
address := RegExReplace(address, "(?<!\w)RTE($|(?!\w)(\.?,?))", "ROUTE")
address := RegExReplace(address, "(?<!\w)SQ($|(?!\w)(\.?,?))", "SQUARE")
address := RegExReplace(address, "(?<!\w)TRL($|(?!\w)(\.?,?))", "TRAIL")
address := RegExReplace(address, "(?<!\w)PLZ($|(?!\w)(\.?,?))", "PLAZA")
address := RegExReplace(address, "(?<!\w)STN($|(?!\w)(\.?,?))", "STATION")
address := RegExReplace(address, "(?<!\w)POST OFFICE BOX($|(?!\w)(\.?,?))", "PO BOX")
StringLower, lower_address, address, T

;---------------Fix any capital case discrepancies---------------
lower_address := RegExReplace(lower_address, "Po Box", "PO Box") 
lower_address := RegExReplace(lower_address, "P.o. Box", "PO Box") 
lower_address := RegExReplace(lower_address, "Po-box", "PO Box")
lower_address := RegExReplace(lower_address, "Nw($|[^A-Za-z])", "NW")
lower_address := RegExReplace(lower_address, "Ne($|[^A-Za-z])", "NE") 
lower_address := RegExReplace(lower_address, "Sw($|[^A-Za-z])", "SW")
lower_address := RegExReplace(lower_address, "Se($|[^A-Za-z])", "SE")
lower_address := RegExReplace(lower_address, "(?<=\d)St", "st")
lower_address := RegExReplace(lower_address, "(?<=\d)Nd", "nd")
lower_address := RegExReplace(lower_address, "(?<=\d)Rd", "rd")
lower_address := RegExReplace(lower_address, "(?<=\d)Th", "th")
lower_address := RegExReplace(lower_address, "(?<!\w)Rr ", "RR #")
lower_address := RegExReplace(lower_address, "(?<!\w)R.r. ", "RR #")
lower_addresss := RegExReplace(lower_address, "(?<!\w)Rpo(?!\w)", "RPO")

;---------------Replace typed-out NESW directions with single letter---------------
lower_address := RegExReplace(lower_address, "North(?!\w)", "N")
lower_address := RegExReplace(lower_address, "East(?!\w)", "E")
lower_address := RegExReplace(lower_address, "South(?!\w)", "S")
lower_address := RegExReplace(lower_address, "West(?!\w)", "W")

Send, !{Tab}
Sleep 100 
Send, !r
Sleep 200
Send, %lower_address%
WinActivate Mozilla Firefox$
Sleep 100
Send, ^f
Sleep 300 
Send, Postal code/Zip code:{Escape}{Control Down}{Shift Down}{Right 2}{Shift Up}{Control Up}
Send, ^c
Sleep 100
Send, {PgUp}
postal_code_line := A_Clipboard
postal_code_cut := RegExMatch(postal_code_line, ":")
postal_code := SubStr(postal_code_line, postal_code_cut+2)
Send, !{Tab}
Sleep 100
Send, {Tab}%postal_code%{Tab}
Sleep 500
Send, !r{Right}
Sleep 300
Pause
Sleep 100
Send, !ac
Pause ;this is here so that if the commercial type needs the producer entered you can do so
Send, {Enter}
Sleep 1000 
WinGetTitle, window_title, A
window_title_cut := RegExMatch(window_title, "- ")
client_name := SubStr(window_title, window_title_cut+2)
Send, %client_name%
Sleep 300
Send, ^s
Sleep 400
Click, 216 201 ;this is supposed to click the 'x' underneath the open contact in Epic
Return

^+4:: ;testing key

;TEST THE REGEX ONCE FUNCTIONAL ON BURNPRE-01

Send, ^f
Sleep 300
Send, Address{Escape}+{Down}^c
Sleep 100
address_untrimmed := A_Clipboard
left_trim := RegExMatch(address_untrimmed, "Address:")
right_trim := RegExMatch(address_untrimmed, "City:")

address := SubStr(address_untrimmed, left_trim+10, right_trim-15)
;address := "24 NORMANDY ST EAST"
;replacements for address abbreviations
address := RegExReplace(address, "(?<!\w)AVE?($|(?!\w)(\.?,?))", "AVENUE") 
address := RegExReplace(address, "(?<!\w)RD($|(?!\w)(\.?,?))", "ROAD") 
address := RegExReplace(address, "(?<!\w)NO($|(?!\w)(\.?,?))", "#") 
address := RegExReplace(address, "(?<!\w)BLVD($|(?!\w)(\.?,?))", "BOULEVARD") 
address := RegExReplace(address, "(?<!\w)HWY($|(?!\w)(\.?,?))", "HIGHWAY") 
address := RegExReplace(address, "(?<!\w)STR?($|(?!\w)(\.?,?))", "STREET") 
address := RegExReplace(address, "(?<!\w)DR($|(?!\w)(\.?,?))", "DRIVE") 
address := RegExReplace(address, "(?<!\w)RANGE ROAD($|(?!\w)(\.?,?))", "RR") 
address := RegExReplace(address, "(?<!\w)CRES($|(?!\w)(\.?,?))", "CRESCENT") 
address := RegExReplace(address, "(?<!\w)CONC\.?($|(?!\w)(\.?,?))", "CONCESSION") 
address := RegExReplace(address, "(?<!\w)CR?T($|(?!\w)(\.?,?))", "COURT") 
address := RegExReplace(address, "(?<!\w)JCT($|(?!\w)(\.?,?))", "JUNCTION") 
address := RegExReplace(address, "(?<!\w)LN($|(?!\w)(\.?,?))", "LANE")
address := RegExReplace(address, "(?<!\w)PKWY($|(?!\w)(\.?,?))", "PARKWAY")
address := RegExReplace(address, "(?<!\w)PL($|(?!\w)(\.?,?))", "PLACE")
address := RegExReplace(address, "(?<!\w)RTE($|(?!\w)(\.?,?))", "ROUTE")
address := RegExReplace(address, "(?<!\w)SQ($|(?!\w)(\.?,?))", "SQUARE")
address := RegExReplace(address, "(?<!\w)TRL($|(?!\w)(\.?,?))", "TRAIL")
address := RegExReplace(address, "(?<!\w)PLZ($|(?!\w)(\.?,?))", "PLAZA")
address := RegExReplace(address, "(?<!\w)STN($|(?!\w)(\.?,?))", "STATION")
address := RegExReplace(address, "(?<!\w)POST OFFICE BOX($|(?!\w)(\.?,?))", "PO BOX")
StringLower, lower_address, address, T

;---------------Fix any capital case discrepancies---------------
lower_address := RegExReplace(lower_address, "Po Box", "PO Box") 
lower_address := RegExReplace(lower_address, "P.o. Box", "PO Box") 
lower_address := RegExReplace(lower_address, "Po-box", "PO Box")
lower_address := RegExReplace(lower_address, "Nw($|[^A-Za-z])", "NW")
lower_address := RegExReplace(lower_address, "Ne($|[^A-Za-z])", "NE") 
lower_address := RegExReplace(lower_address, "Sw($|[^A-Za-z])", "SW")
lower_address := RegExReplace(lower_address, "Se($|[^A-Za-z])", "SE")
lower_address := RegExReplace(lower_address, "(?<=\d)St", "st")
lower_address := RegExReplace(lower_address, "(?<=\d)Nd", "nd")
lower_address := RegExReplace(lower_address, "(?<=\d)Rd", "rd")
lower_address := RegExReplace(lower_address, "(?<=\d)Th", "th")
lower_address := RegExReplace(lower_address, "(?<!\w)Rr ?", "RR #")
lower_address := RegExReplace(lower_address, "(?<!\w)R.r. ", "RR #")

;---------------Replace typed-out NESW directions with single-letter---------------
lower_address := RegExReplace(lower_address, "North(?!\w)", "N")
lower_address := RegExReplace(lower_address, "East(?!\w)", "E")
lower_address := RegExReplace(lower_address, "South(?!\w)", "S")
lower_address := RegExReplace(lower_address, "West(?!\w)", "W")
MsgBox %lower_address%
Return

^+5::
setTitleMatchMode RegEx
;WinActivate [^(Inbox|Prospect List|OneDrive)] - [\w]+
WinActivate (?<!Inbox)(?<!Prospect)(?<!OneDrive) - (?!Landon)(?!Excel)(?!3CX)(?!Notepad)(?!Google)(?!Mozilla)[a-zA-Z]+
WinGetActiveTitle, currentwindow
MsgBox %currentwindow%
Return

^+1::
Click, 216 201
Return