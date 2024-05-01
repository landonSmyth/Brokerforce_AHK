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
-This works assuming that the RR is 0001, you must check that it is before finalizing
-This also depends upon the rate sheet being the 2022 rate sheet and the only rate sheet
    ->If there are two rate sheets for a client, you need to fix the revenue in both of them
-This depends on the most recent rate sheet being attached with the "SPREADSHEET" folder in Epic.
Sometimes other people forget to do so, so check that the premium indication in the bottom
is ALWAYS 2022 renewal premium, if it is 2021 you need to go fetch the correct rate sheet
    ->You will know for sure if it is a 2021 sheet because the Ctrl+F for "2022 Renewal Premium" will fail
-This script was desgined for Firefox as the browser used to search CRA, the logic for copying the 
revenue from the website without using the mouse might not line up correctly with other browsers
-If CRA ever change the structure of the advanced search, the Tabs may be thrown off and would need updated
*/

;---------------Hotkey for single construction rate sheets---------------
^+0:: 
setTitleMatchMode RegEx
Send, ^f2023 Ren!f
Sleep 200 ;This sleep time must be here to allow Excel to find the cell before moving forward 
Send, {Escape}
Send, 2024 Renewal Premium: ${Enter}
Send, !hzefdg
Sleep 500 
Send, k3{Enter}
Send, %A_YYYY%-%A_MM%-%A_DD%{Enter}
Send, !hzefdg
Sleep 500 
Send, b7{Enter}{F2}{Up}
Sleep 100
Send, {Shift Down}{Right 9}{Shift Up}^c
Send, {Escape}!hzefdg
Sleep 500 
Send, b6{Enter}
WinActivate Mozilla Firefox$
Sleep 200
Send, {Tab}{Enter}
Sleep 200
Send, {Tab 2}^v{Tab 2}0001 
Pause
Send, {Enter}
WinWaitActive ^T3010
Sleep 500
Send, ^f
Sleep 300
Send, Total revenue
Pause ;this is here in case their revenue is not displayed. If it is not, you can reload and go from there
Send, {Escape}{Control Down}{Shift Down}{Right 5}{Shift Up}{Control Up}^c
Sleep 50
Revenue_full := A_Clipboard
Cut_position := RegExMatch(Revenue_full, ":")
Revenue_trimmed := SubStr(Revenue_full, Cut_position+3)
A_Clipboard := ""
Send, {PgUp 3}
WinActivate Excel$
Sleep 100
Send, {F2}{Shift Down}{Up}{Shift Up}%Revenue_trimmed%{Tab}
Pause
Send, !{F4}
Sleep 100
WinActivate Mozilla Firefox$
Sleep 300
Send, ^l
Sleep 100
Send, https://apps.cra-arc.gc.ca/ebci/hacc/srch/pub/dsplyAdvncdSrch
Sleep 400
Send, {Enter}
WinActivate ^\d\d? - [A-Z]{3}
Sleep 100
Send, {Down}
Return

^+8:: ;hotkey to send 2023 Renewal Premium for a rate sheet that doesnt fit the main hotkey	
Send, 2024 Renewal Premium: ${Enter}
Return

^+7:: ;hotkey to send the current date
Send, %A_MM%-%A_DD%{Enter}
Return

;---------------hotkey for only updating the revenue on the sheet---------------
^+9:: 
setTitleMatchMode RegEx
Send, !hzefdg
Sleep 500 
Send, k3{Enter}
Send, %A_YYYY%-%A_MM%-%A_DD%{Enter}
Send, !hzefdg
Sleep 350 
Send, b7{Enter}{F2}{Up}
Sleep 100
Send, {Shift Down}{Right 9}{Shift Up}^c
Send, {Escape}!hzefdg
Sleep 350 
Send, b6{Enter}
WinActivate Mozilla Firefox$
Sleep 200
Send, {Tab}{Enter}
Sleep 200
Send, {Tab 2}^v{Tab 2}0001 
Pause
Send, {Enter}
Sleep 1700
Send, ^f
Sleep 300
Send, Total revenue
Pause ;this is here in case their revenue is not displayed. If it is not, you can reload and go from there
Send, {Escape}{Control Down}{Shift Down}{Right 5}{Shift Up}{Control Up}^c
Sleep 50
Revenue_full := A_Clipboard
Cut_position := RegExMatch(Revenue_full, ":")
Revenue_trimmed := SubStr(Revenue_full, Cut_position+3)
A_Clipboard := ""
Send, {PgUp 3}
WinActivate Excel$
Sleep 100
Send, {F2}{Shift Down}{Up}{Shift Up}%Revenue_trimmed%{Tab}
Pause
Send, !{F4}
Sleep 100
WinActivate Mozilla Firefox$
Sleep 300
Send, ^l
Sleep 100
Send, https://apps.cra-arc.gc.ca/ebci/hacc/srch/pub/dsplyAdvncdSrch
Sleep 350
Send, {Enter}
Sleep 100
WinActivate ^\d\d? - [A-Z]{3}
Sleep 150
Send, {Down}
Return

;---------------Hotkey to reset the CRA search and close the rate sheet separately (if a reload of the main key is needed prior)--------
^+2:: 
setTitleMatchMode RegEx
Send, !{F4}
Sleep 100
WinActivate Mozilla Firefox$
Sleep 300
Send, ^l
Sleep 100
Send, https://apps.cra-arc.gc.ca/ebci/hacc/srch/pub/dsplyAdvncdSrch
Sleep 300
Send, {Enter}
WinActivate ^\d\d? - [A-Z]{3}
Return

/*
This hotkey is more reliable because it enters the RR as it appears, but is slightly slower because two window swaps are required
for entering the BN and RR separately. Trying to do them at the same time results in the clipboard reference overwriting the BN number.
if ever giving this script to someone else, have them use this key if they would prefer stability for all possible rate sheets. 
*/
^+4:: 
setTitleMatchMode RegEx
Send, !hzefdg
Sleep 350 
Send, b7{Enter}{F2}{Up}
Sleep 100
Send, {Shift Down}{Right 9}{Shift Up}^c
Sleep 50
BN := A_Clipboard
WinActivate Mozilla Firefox$
Sleep 200
Send, {Tab}{Enter}
Sleep 200
Send, {Tab 2}%BN%
WinActivate Excel$
Sleep 100
Send, {Down 2}{Shift Down}{Left 4}{Shift Up}^c
Sleep 50
RR := A_Clipboard
Send, {Escape}!hzefdg
Sleep 350 
Send, b6{Enter}
WinActivate Mozilla Firefox$
Send, {Tab 2}%RR%
Pause
Send, {Enter}
Sleep 1700
Send, ^f
Sleep 300
Send, Total revenue
Pause ;this is here in case their revenue is not displayed. If it is not, you can reload and go from there
Send, {Escape}{Control Down}{Shift Down}{Right 5}{Shift Up}{Control Up}^c
Sleep 50
Revenue_full := A_Clipboard
Cut_position := RegExMatch(Revenue_full, ":")
Revenue_trimmed := SubStr(Revenue_full, Cut_position+3)
A_Clipboard := ""
Send, {PgUp 3}
WinActivate Excel$
Sleep 100
Send, {F2}{Shift Down}{Up}{Shift Up}%Revenue_trimmed%{Tab}
Pause
Send, !{F4}
Sleep 100
WinActivate Mozilla Firefox$
Sleep 300
Send, !b
Sleep 200
Send, {Down 3}{Enter}
Sleep 100
Send, {Enter}
WinActivate ^\d\d? - [A-Z]{3}
Return

/*
^+1:: Rate sheet for Quebec
ctrl f 2022 premium is the same
date field is located at M3
crc number is at B8
Revenue is at B7

^+1:: Rate sheet for Multi-construction
ctrl f 2022 premium is the same
date field is located at L3
crc number is at B8
Revenue is at B7
*/