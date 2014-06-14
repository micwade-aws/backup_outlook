; File:    backup.ahk
; Author:  Michael J. Wade
; Purpose: Exports a backup of your Outlook inbox to a pst file using the date as a filename
;          Tweak this script however you need.  It's heavily commented and AHK scripting is EASY.

; Stuff from the AHK template:
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Open outlook
Run Outlook.exe
Sleep 10000   ; Waits for 10 seconds - is that enough time?

; Keystrokes needed to get into the export functionality
; Alt+F, then Alt+O, then Alt+I
Send !{F}!{O}!{I}


; Up arrow 3 times to hit the "Export to file" option
Send {Up}{Up}{Up}{Enter}
Sleep 1000

; Down to hit Outlook Data File (.pst)
Send {Down}{Enter}
Sleep 1000

; Default is inbox and it's subfolders.
; If you want to capture everything, uncomment next line and add/remove "ups"
; Send {Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}{Up}
Send {Enter}
Sleep 1000

; Backspace the default filename and replace with this location:
Send {Backspace}
FormatTime, CurrentDateTime,, MM-dd-yyyy ; It will look like 9-1-2005
Send +%A_ScriptDir%\email-%CurrentDateTime%.pst
Send {Enter}

; Password Screen?  No password, just enter.
Send {Enter}
