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
-This works only assuming the the most recent rate sheet is attached to the "SPREADSHEET" folder in Epic. If it is not, the script will add
the next most recent rate sheet in that folder, likely one from the year before. Double check to ensure consistency
-NOTE: you must be paying attention to the results from the search in Epic. If they have two rate sheets like an "A" and "B"
sheet for two locations, you must manually attach the second one and rename to "A" and "B" accordingly.
-You must manually update the folder path to be the desired month. 
*/

RShift::
Send, !{F4}
Return

^+0:: ;this key is calls itself so that you dont have to press the button every time. Non-recursive copy bound to Ctrl+Shft+9
recursive_key(){
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
StringUpper, client_code, client_code
Send, %client_code%{Enter}!aa
Pause
WinGetActiveTitle, client_window
if(InStr(client_window, "(") != 0 or InStr(client_window, ")") != 0){
	client_window := SubStr(client_window, 1, 13)
}
Send, {Tab 2}f{Tab 2}spr{Enter 2}
WinWaitActive Excel$
Sleep 1000 ;this likely would need tuning depending on the system or the loading time of Excel
Send, ^f2023 Ren!f
Sleep 200 ;This sleep time must be here to allow Excel to find the cell before moving forward 
Send, {Escape}
Send, 2024 Renewal Premium: ${Enter}
Send, !hzefdg
Sleep 500 
Send, k3{Enter}
Send, %A_YYYY%-%A_MM%-%A_DD%{Enter}
Sleep 500
Send, !fa
Sleep 500
Send o
Sleep 500
Send, %client_code% R 24
Sleep 50
Send {F4}^a
Sleep 100
Send C:\Users\LandonSmyth\Brokerforce Insurance Inc\BrokerForce shared - Documents\Sanctuary Plus\Rate Sheets\10 - OCTOBER
Sleep 100
Send, {Enter}
A_Clipboard := client_code . " R 24"
Pause
Send, !s
Sleep 300
Send, !{F4}
WinWaitActive %client_window%
Sleep 1000
Send, !{F4}
while WinExist(client_window){
	if WinActive(client_window) or WinActive("^Closing"){
		Send, !{F4}
		Sleep 350
	}
}
WinWaitActive ^Account Locate
recursive_key()
}
recursive_key()
Return

^+9::
setTitleMatchMode RegEx
InputBox, client_code, Client Code, Enter client code
Send, %client_code%{Enter}!aa
Pause
WinGetActiveTitle, client_window
if(InStr(client_window, "(") != 0 or InStr(client_window, ")") != 0){
	client_window := SubStr(client_window, 1, 13)
}
Send, {Tab 2}f{Tab 2}spr{Enter 2}
WinWaitActive Excel$
Sleep 1000 ;this likely would need tuning depending on the system or the loading time of Excel
Send, !fa
Sleep 500
Send o
Sleep 500
Send, %client_code% R 24
Sleep 50
Send {F4}^a
Sleep 100
Send C:\Users\LandonSmyth\Brokerforce Insurance Inc\BrokerForce shared - Documents\Sanctuary Plus\Rate Sheets\12 - DECEMBER
Sleep 100
Send, {Enter}
Pause
Send, !s
Sleep 300
Send, !{F4}
WinWaitActive %client_window%
Sleep 1000
Send, !{F4}
Return
