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

MButton::Send, !{Left}

;---------------Hotkey to lookup client and search them in CRA---------------
^+0:: ;start by looking up the client by their name and pressing enter
setTitleMatchMode RegEx
Send, !ad
Pause
WinGetActiveTitle, client_window
Send, ^c
Sleep 100
client_name := A_Clipboard
IfWinExist, ^List of
{
	WinActivate ^List of
}
else {
	MsgBox Search Window not open, open it and unpause
	Sleep 500
	Pause
}
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
WinActivate %client_window%
Return

^+7::
;---------------Hotkey to update address based on CRA information---------------
/*
before starting the script, ensure the client's name is in your clipboard. Either it is already in your clipboard from
searching in CRA, or you need to copy the updated legal name from CRA before starting
*/

setTitleMatchMode RegEx
WinActivate (^List of|^T3010|Firefox)
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

WinActivate (?<!Inbox)(?<!Prospect)(?<!OneDrive) - (?!Landon)(?!Excel)(?!3CX)(?!Notepad)[a-zA-Z]+
Sleep 100 
Send, !r
Sleep 200
Send, %lower_address%
WinActivate (^List of|^T3010|Firefox)
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
Return

;---------------Hotkey to Reset the CRA lookup and close client---------------
^+2::
setTitleMatchMode RegEx
WinActivate (?<!Inbox)(?<!Prospect)(?<!OneDrive) - (?!Landon)(?!Excel)(?!3CX)(?!Notepad)(?!Google)(?!Mozilla)[a-zA-Z]+
Send, !{F4}
WinActivate (^List of|^T3010|Firefox)
Sleep 300
Send, ^l
Sleep 100
Send, https://apps.cra-arc.gc.ca/ebci/hacc/srch/pub/dsplyAdvncdSrch
Sleep 300
Send, {Enter}
WinActivate Account Locate
Sleep 100
Send, !c{Tab 6}
Return

;---------------Hotkey to type out the client name in title case---------------
^+9::
WinGetTitle, window_title, A
window_title_cut := RegExMatch(window_title, "- ")
client_name := SubStr(window_title, window_title_cut+2)
StringLower, client_name, client_name, T
Send, %client_name%
Return











