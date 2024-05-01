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

;TEST ONCE FULLY WORKING ON FIRST21OCT, this is a client that has emails in both tabs and websites to cut out

^+0::
setTitleMatchMode RegEx
Click, 121 592  
Sleep 100
Click, 121 592 Right 
Sleep 100
Click 185 629 ;update this
Sleep 100
Click, 121 592 Right 
Sleep 100
Click, 174 604
Sleep 100
info := A_Clipboard

left_trim := RegExMatch(info, "-\d\d\d\d")
right_trim := RegExMatch(info, "Contact via")
emails := SubStr(info, left_trim+8, -20)
A_Clipboard := emails
MsgBox %A_Clipboard%

Return


^+1::
/*
What is needed:
-dont paste anything if an email is not ther
-if an email is in there, dont accidentally paste 

Test on these clients once working
BIBL-12
ARMEN02
*/
info:= "
(
Fax: (416) 441-3379 
www.jesusfirstassembly.com
)"
trim := RegExMatch(info, "[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+")
;MsgBox Trim is %trim%
email_untrimmed := SubStr(info, trim)
MsgBox %email_untrimmed%
email := RegExReplace(email_untrimmed, "(https:\/\/www\.|http:\/\/www\.|https:\/\/|http:\/\/)?[a-zA-Z0-9]{2,}(\.[a-zA-Z0-9]{2,})(\.[a-zA-Z0-9]{2,})?$", "")
MsgBox %email%
Return

;officeadmin@openarmsparrsboro.ca
;^[A-Z0-9+_.-]+@[A-Z0-9.-]+$

;Fax: (416) 441-3379 
;www.jesusfirstassembly.com

;Residence: (416) 223-7982 
;Fax: (416) 279-1060 
;Business: (905) 882-9996 

;Business: (905) 770-8828 231 
;Fax: (905) 770-8851 
;Jackson.Wong@brokerteam.ca

;erica.tibbo@gmail.com
;www.firstbaptistporthope.com

