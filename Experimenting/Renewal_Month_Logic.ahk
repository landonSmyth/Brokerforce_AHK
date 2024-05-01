

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Del::ExitApp

`::
Pause -1
Return

PgUp::
reload
Return

^+0::
Policy_term := "July 1, 2023 - July 3, 2024" ;policy term taken as input with an input box
Cut_position := RegExMatch(Policy_term, "-") ;Find the integer position at which the effective date ends, this is variable because of month length
Renewal_date := SubStr(Policy_term, 1, Cut_position-2) ;Store the renewal date, meaning the whole string prior to the " -"
month_array := ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
;Month array is used for the month + 1 functionality for the three-month pay plan 
month_to_number := {"January":1, "February":2, "March":3, "April":4, "May":5, "June":6, "July":7, "August":8, "September":9, "October":10, "November":11, "December":12}
;month to number is used for storing the related month number for indexing month_array when entering the three-pay plan
month_cut := RegExMatch(Renewal_date, " \d,") ;Find the integer position of where the month string is 
month := SubStr(Renewal_date, 1, month_cut-1) ;Store the month in a variable 
Send, %month%
Month_number := month_to_number[month] ;Store the related month number for indexing later on 
Send, %Month_number%
Return

^+1::
Send, ^f
Sleep 300
Send, Total revenue
Pause
Send, {Escape}{Control Down}{Shift Down}{Right 5}{Shift Up}{Control Up}^c
Sleep 50
Revenue_full := A_Clipboard
Cut_position := RegExMatch(Revenue_full, ":")
MsgBox, %Cut_position%
Revenue_trimmed := SubStr(Revenue_full, Cut_position+3)
MsgBox, %Revenue_trimmed%
A_Clipboard := ""
Return
;pseudocode for highlighting revenue without mouse:
;After ctrl f Total Revenue:
;Escape, Control & Shift Down {Right} * 5, ctrl C
;use sub string regex to trim the string down to the revenue
;continue as normal
;8677215240001 is $1m+ client
;1280095940001 is a 6-figure client
;106923782 is a 5-figure client
