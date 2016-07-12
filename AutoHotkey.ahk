; IMPORTANT INFO ABOUT GETTING STARTED: Lines that start with a
; semicolon, such as this one, are comments.  They are not executed.

; This script has a special filename and path because it is automatically
; launched when you run the program directly.  Also, any text file whose
; name ends in .ahk is associated with the program, which means that it
; can be launched simply by double-clicking it.  You can have as many .ahk
; files as you want, located in any folder.  You can also run more than
; one .ahk file simultaneously and each will get its own tray icon.

; SAMPLE HOTKEYS: Below are two sample hotkeys.  The first is Win+Z and it
; launches a web site in the default browser.  The second is Control+Alt+N
; and it launches a new Notepad window (or activates an existing one).  To
; try out these hotkeys, run AutoHotkey again, which will load this file.

;#z::Run www.autohotkey.com

;^!n::
;IfWinExist Untitled - Notepad
;	WinActivate
;else
;	Run Notepad
;return



; Note: From now on whenever you run AutoHotkey directly, this script
; will be loaded.  So feel free to customize it to suit your needs.

; Please read the QUICK-START TUTORIAL near the top of the help file.
; It explains how to perform common automation tasks such as sending
; keystrokes and mouse clicks.  It also explains more about hotkeys.
;MsgBox %A_TitleMatchMode%

CapsLock::WinMinimize,A
+CapsLock::CapsLock
; $ causes the F12 press not to be sent also
$F12::SendInput,{Volume_Mute}
$F11::SendInput,{Volume_Up 5}
$F10::SendInput,{Volume_Down 5}
$F9::SendInput,{Media_Next}
;$!F9::SendInput,{Media_Prev}
$^!F9::SendInput,{Media_Play_Pause}
^+F12:: SendInput, {F12}
^+F11:: SendInput, {F11}
^+F10:: SendInput, {F10}
^+F9:: SendInput, {F9}
^+!o::
;~ WinActivate Inbox ahk_class rctrl_renwnd32
; get the number of monitors and store it in MonitorCount
SysGet, MonitorCount, MonitorCount
SysGet, MonitorPrimary, MonitorPrimary
SysGet, PrimaryMonWorkArea, MonitorWorkArea, %MonitorPrimary%



;~ if (MonitorCount >= 2)
;~ {
	;~ if (MonitorPrimary = 1)
	;~ {
		;~ global SecondMonitor := 2
	;~ } else {
		;~ global SecondMonitor := 1
	;~ }
;~ }

global SecondMonitor := 2
global ThirdMonitor := 3

SysGet, SecondMonWorkArea, MonitorWorkArea, %SecondMonitor%		
SysGet, ThirdMonWorkArea, MonitorWorkArea, %ThirdMonitor%	
;MsgBox, Left: %ThirdMonWorkAreaLeft% -- Top: %ThirdMonWorkAreaTop% -- Right: %ThirdMonWorkAreaRight% -- Bottom %ThirdMonWorkAreaBottom%.
CommunicatorWidth = 402
TaskBarWidth = %PrimaryMonWorkAreaLeft%
;Debugging
;Difference := (SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
;ListVars

;IfWinNotExist, Inbox - Microsoft Outlook ahk_class rctrl_renwnd32 
IfWinNotExist, Inbox ahk_class rctrl_renwnd32
{
	Run, "C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE"
	WinWait, Inbox ahk_class rctrl_renwnd32	
}
WinActivate, Inbox ahk_class rctrl_renwnd32
if (MonitorCount = 1){
	WinMove, Inbox ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
} else if (MonitorCount = 2){
	WinMove, Inbox ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
} else if (MonitorCount = 3){
	WinMove, Inbox ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

IfWinNotExist, Archive ahk_class rctrl_renwnd32
{
	Run, "C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE"  /select Outlook:\\NewArchive\Archive
	WinWait, Archive ahk_class rctrl_renwnd32
}
WinActivate, Archive ahk_class rctrl_renwnd32

if (MonitorCount = 1) {
	WinMove, Archive ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
} else if (MonitorCount = 2){
	WinMove, Archive ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
} else if (MonitorCount = 3){
	WinMove, Archive ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft)*5/6,  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

IfWinNotExist, Calendar ahk_class rctrl_renwnd32
{
	Run, "C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE"  /select outlook:calendar	
	WinWait, Calendar ahk_class rctrl_renwnd32
}
WinActivate, Calendar ahk_class rctrl_renwnd32

if (MonitorCount = 1) {
	WinMove, Calendar ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft),  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
} else if (MonitorCount = 2){
	WinMove, Calendar ahk_class rctrl_renwnd32, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, Abs(SecondMonWorkAreaRight - SecondMonWorkAreaLeft - CommunicatorWidth), Abs(SecondMonWorkAreaBottom - SecondMonWorkAreaTop) 
} else if (MonitorCount = 3){
	WinMove, Calendar ahk_class rctrl_renwnd32, , %ThirdMonWorkAreaLeft%, %ThirdMonWorkAreaTop%, Abs(ThirdMonWorkAreaRight - ThirdMonWorkAreaLeft)*5/6,  Abs(ThirdMonWorkAreaBottom - ThirdMonWorkAreaTop)
}
;Outlook Tasks window
;~ IfWinNotExist, Tasks ahk_class rctrl_renwnd32
;~ {
	;~ Run, "C:\Program Files\Microsoft Office 15\root\office15\OUTLOOK.EXE"  /select outlook:tasks	
	;~ WinWait, Tasks ahk_class rctrl_renwnd32
;~ }
;~ WinActivate, Tasks ahk_class rctrl_renwnd32

;~ if (MonitorCount = 1) {
	;~ WinMove, Tasks ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft),  Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
;~ } else if (MonitorCount = 2){
	;~ WinMove, Tasks ahk_class rctrl_renwnd32, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, Abs(SecondMonWorkAreaRight - SecondMonWorkAreaLeft - CommunicatorWidth), Abs(SecondMonWorkAreaBottom - SecondMonWorkAreaTop) 
;~ } else if (MonitorCount = 3){
	;~ WinMove, Tasks ahk_class rctrl_renwnd32, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, Abs(SecondMonWorkAreaRight - SecondMonWorkAreaLeft - CommunicatorWidth), Abs(SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
;~ }

IfWinNotExist, Skype for Business  ahk_class CommunicatorMainWindowClass
{
	Run, lync
	WinWait, Skype for Business  ahk_class CommunicatorMainWindowClass
}
WinActivate, Skype for Business  ahk_class CommunicatorMainWindowClass
if (MonitorCount >= 2){
	WinMove, Skype for Business  ahk_class CommunicatorMainWindowClass, , %SecondMonWorkAreaLeft%, %SecondMonWorkAreaTop%, %CommunicatorWidth%, Abs(SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
} else {
	WinMove, Skype for Business  ahk_class CommunicatorMainWindowClass, , (PrimaryMonWorkAreaRight - CommunicatorWidth), %PrimaryMonWorkAreaTop%, %CommunicatorWidth%, Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

WinActivate, Inbox ahk_class rctrl_renwnd32
return

^+!g::
SysGet, MonitorCount, MonitorCount
SysGet, MonitorPrimary, MonitorPrimary
SysGet, PrimaryMonWorkArea, MonitorWorkArea, %MonitorPrimary%
if (MonitorCount >= 2)
{
	if (MonitorPrimary = 1)
	{
		global SecondMonitor = 2
	} else {
		global SecondMonitor = 1
	}
}
SysGet, SecondMonWorkArea, MonitorWorkArea, %SecondMonitor%		
CommunicatorWidth = 385
TaskBarWidth = %PrimaryMonWorkAreaLeft%
;Debugging
;Difference := (SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
;ListVars

SetTitleMatchMode, 2
WinActivate, ahk_class Chrome_WidgetWin_1
if (MonitorCount >= 2){
	WinMove, ahk_class Chrome_WidgetWin_1, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, Abs(SecondMonWorkAreaRight - SecondMonWorkAreaLeft) - CommunicatorWidth, Abs(SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
} else {
	WinMove, ahk_class Chrome_WidgetWin_1, , TaskBarWidth, %PrimaryMonWorkAreaTop%, Abs(PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft), Abs(PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

SetTitleMatchMode, 1
return

;Hotstrings (text expansion)
:*:gmh::george.murga@honeywell.com
:*:gmg::george.murga@gmail.com
::urlenc::
	oldClip := Clipboard
	Clipboard := MakeNiceURL()
    SendInput, ^v
	ClipWait
	Clipboard := oldClip
return
::urldec::
	oldClip := Clipboard
	Clipboard := MakeTitleFromNiceURL()
    SendInput, ^v
	ClipWait
	Clipboard := oldClip
return

MakeNiceURL()
{
	NiceUrl := Trim(Clipboard)
	While InStr(NiceUrl, "  ") <> 0
	{
		StringReplace, NiceUrl, NiceUrl, %A_Space%%A_Space%, %A_Space%, All
	}
	StringReplace, NiceUrl, NiceUrl, :, , All
	StringReplace, NiceUrl, NiceUrl, `;, , All
	StringReplace, NiceUrl, NiceUrl, `,, , All
	StringReplace, NiceUrl, NiceUrl, !, , All
	StringReplace, NiceUrl, NiceUrl, ., , All
	StringReplace, NiceUrl, NiceUrl, /, -, All
	StringReplace, NiceUrl, NiceUrl, \, -, All
	StringReplace, NiceUrl, NiceUrl, #, , All
	StringReplace, NiceUrl, NiceUrl, (, , All
	StringReplace, NiceUrl, NiceUrl, ), , All
	StringReplace, NiceUrl, NiceUrl, [, , All
	StringReplace, NiceUrl, NiceUrl, ], , All
	StringReplace, NiceUrl, NiceUrl, `{, , All
	StringReplace, NiceUrl, NiceUrl, `}, , All
	StringReplace, NiceUrl, NiceUrl, ', , All
	StringReplace, NiceUrl, NiceUrl, `", , All
	StringReplace, NiceUrl, NiceUrl,  %A_Space%, -, All
	While InStr(NiceUrl, "--") <> 0
	{
		StringReplace, NiceUrl, NiceUrl, --, -, All
	}
	return NiceUrl
}

MakeTitleFromNiceURL()
{
	NiceUrl := Trim(Clipboard)
	StringReplace, NiceUrl, NiceUrl,  -, %A_Space%, All
	While InStr(NiceUrl, "  ") <> 0
	{
		StringReplace, NiceUrl, NiceUrl, %A_Space%%A_Space%, %A_Space%, All
	}
	StringUpper, NiceUrl, NiceUrl, T
	return NiceUrl
}