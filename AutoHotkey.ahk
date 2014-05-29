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
F12::Send,{Volume_Mute}
F11::Send,{Volume_Up 5}
F10::Send,{Volume_Down 5}
^+F12::F12
^+F11::F11
^+F10::F10
^+!o::
;~ WinActivate Inbox ahk_class rctrl_renwnd32
; get the number of monitors and store it in MonitorCount
SysGet, MonitorCount, MonitorCount
sysget, MonitorPrimary, MonitorPrimary
SysGet, PrimaryMonWorkArea, MonitorWorkArea, %MonitorPrimary%
if (MonitorCount = 2)
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

IfWinNotExist, Inbox - Microsoft Outlook ahk_class rctrl_renwnd32
{
	Run, "c:\Program Files (x86)\Microsoft Office\Office12\Outlook.exe"
	WinWait, Inbox ahk_class rctrl_renwnd32	
}
WinActivate, Inbox ahk_class rctrl_renwnd32
WinMove, Inbox ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, (PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft),  (PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)

IfWinNotExist, Archive ahk_class rctrl_renwnd32
{
	Run, "c:\Program Files (x86)\Microsoft Office\Office12\Outlook.exe"  /select Outlook:\\NewArchive\Archive
	WinWait, Archive ahk_class rctrl_renwnd32
}
WinActivate, Archive ahk_class rctrl_renwnd32
WinMove, Archive ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, (PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft),  (PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)

IfWinNotExist, Calendar - Microsoft Outlook ahk_class rctrl_renwnd32
{
	Run, "c:\Program Files (x86)\Microsoft Office\Office12\Outlook.exe"  /select outlook:calendar	
	WinWait, Calendar ahk_class rctrl_renwnd32
}
WinActivate, Calendar ahk_class rctrl_renwnd32

if (MonitorCount = 2){
	WinMove, Calendar ahk_class rctrl_renwnd32, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, (SecondMonWorkAreaRight - SecondMonWorkAreaLeft - CommunicatorWidth), (SecondMonWorkAreaBottom - SecondMonWorkAreaTop) 
} else {
	WinMove, Calendar ahk_class rctrl_renwnd32, , TaskBarWidth, %PrimaryMonWorkAreaTop%, (PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft),  (PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}


IfWinNotExist, Microsoft Lync ahk_class CommunicatorMainWindowClass
{
	Run, Communicator
	WinWait, Microsoft Lync ahk_class CommunicatorMainWindowClass
}
WinActivate, Microsoft Lync ahk_class CommunicatorMainWindowClass
if (MonitorCount = 2){
	WinMove, Microsoft Lync ahk_class CommunicatorMainWindowClass, , %SecondMonWorkAreaLeft%, %SecondMonWorkAreaTop%, %CommunicatorWidth%, (SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
} else {
	WinMove, Microsoft Lync ahk_class CommunicatorMainWindowClass, , (PrimaryMonWorkAreaRight - CommunicatorWidth), %PrimaryMonWorkAreaTop%, %CommunicatorWidth%, (PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

WinActivate, Inbox ahk_class rctrl_renwnd32
return

^+!g::
SysGet, MonitorCount, MonitorCount
SetTitleMatchMode, 2
WinActivate, ahk_class Chrome_WidgetWin_0
;MsgBox %MonitorCount%
if (MonitorCount = 2){
	WinMove, ahk_class Chrome_WidgetWin_0, , (SecondMonWorkAreaLeft + CommunicatorWidth), %SecondMonWorkAreaTop%, (SecondMonWorkAreaRight - SecondMonWorkAreaLeft), (SecondMonWorkAreaBottom - SecondMonWorkAreaTop)
} else {
	WinMove, ahk_class Chrome_WidgetWin_0, , TaskBarWidth, %PrimaryMonWorkAreaTop%, (PrimaryMonWorkAreaRight - PrimaryMonWorkAreaLeft), (PrimaryMonWorkAreaBottom - PrimaryMonWorkAreaTop)
}

SetTitleMatchMode, 1
return