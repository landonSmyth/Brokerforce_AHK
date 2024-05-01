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


;NOTES & DEPENDENCIES
;This works assuming that the RR is 0001, you must check that it is before finalzing
;Later on, implement a snippet to copy the four digits to the right of RR and assign 
;that to a variable for searching the client in CRA
;This also depends upon the rate sheet being the 2022 rate sheet and the only rate sheet
;If there are two rate sheets you need to fix the revenue in both of them
;This depends on the most recent rate sheet being attached with the "SPREADSHEET" folder
;sometimes other people forget to do so, so check that the premium indication in the bottom
;is ALWAYS 2022 renewal premium, if it is 2021 you need to go fetch the correct rate sheet


;pseudocode for highlighting revenue without mouse:
;After ctrl f Total Revenue:
;Escape, Control & Shift Down {Right} * 5, ctrl C
;use sub string regex to trim the string down to the revenue
;continue as normal
;8677215240001 is $1m+ client
;1280095940001 is a 6-figure client
;106923782 is a 5-figure client

^+0:: ;Rate sheet for single construction 
setTitleMatchMode RegEx
Send, ^f2022 Ren!n
Sleep 200 ;This sleep must be here to allow Excel to find the cell before moving forward 
Send, {Escape}
Send, 2023 Renewal Premium: {Enter}
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
WinActivate ^List of
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
Pause
;here you highlight and copy it
Send, {PgUp 3}
WinActivate R 23
Sleep 100
Send, {F2}{Control Down}{Left}{Control Up}^v{Tab}
Pause
Send, !{F4}
Sleep 100
WinActivate ^T3010
Sleep 300
Send, !b
Sleep 200
Send, {Down 3}{Enter}
Sleep 100
Send, {Enter}
WinActivate ^10 - OCT
Return

^+1:: ;Copy of above with broken RR logic
setTitleMatchMode RegEx
Send, ^f2022 Ren!n
Sleep 200 ;This sleep must be here to allow Excel to find the cell before moving forward 
Send, {Escape}
Send, 2023 Renewal Premium: {Enter}
Send, !hzefdg
Sleep 500 
Send, k3{Enter}
Send, %A_YYYY%-%A_MM%-%A_DD%{Enter}
Send, !hzefdg
Sleep 500 
Send, b7{Enter}{F2}{Up}
Sleep 100
Send, {Shift Down}{Right 9}{Shift Up}^c
Sleep 50
BN := A_Clipboard
Send, {Down 2}{Shift Down}{Left 4}{Shift Up}^c
Sleep 50
RR := A_Clipboard
Send, {Escape}!hzefdg
Sleep 500 
Send, b6{Enter}
WinActivate ^List of
Sleep 200
Send, {Tab}{Enter}
Sleep 200
Send, {Tab 2}%BN%{Tab 2}%RR% ;problem is that they are both reassigned when new thing copied
Pause
Send, {Enter}
Sleep 1700
Send, ^f
Sleep 300
Send, Total revenue
Pause
;here you highlight and copy it
Send, {PgUp 3}
WinActivate R 23
Sleep 100
Send, {F2}{Control Down}{Left}{Control Up}^v{Tab}
Pause
Send, !{F4}
Sleep 100
WinActivate ^T3010
Sleep 300
Send, !b
Sleep 200
Send, {Down 3}{Enter}
Sleep 100
Send, {Enter}
WinActivate ^10 - OCT
Return

;^+1:: ;Rate sheet for Quebec
;ctrl f 2022 premium is the same
;date field is located at M3
;crc number is at B8
;Revenue is at B7

;^+1:: ;Rate sheet for Multi-construction
;ctrl f 2022 premium is the same
;date field is located at L3
;crc number is at B8
;Revenue is at B7