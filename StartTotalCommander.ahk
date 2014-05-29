#NoEnv
Run, TOTALCMD64.EXE, , UseErrorLevel
If A_LastError <> 0
{
    MsgBox, 48, Something went wrong..., Please run this program in the same folder as TOTALCMD64.EXE
    ExitApp
}
WinWait, ahk_class TNASTYNAGSCREEN
WinActivate

;Winget, WInfo, ControlList
IfWinExist, ahk_class TNASTYNAGSCREEN 
{
    WinActivate
    DetectHiddenText, Off
    SetTitleMatchMode, Fast
    ;WinGetText, WTextSlow, ahk_class TNASTYNAGSCREEN 
    loop {
        ControlGetText, ButtonPress, Window4, ahk_class TNASTYNAGSCREEN
        if ButtonPress <>
            break
        }
    ;Button2Press := "Button" . ButtonPress   
    ;ListVars
    ;MsgBox, %ButtonPress% ;%WTextSlow%

    WinWait, ahk_class TNASTYNAGSCREEN
    WinActivate
    ;ControlClick, %Button2Press%, ahk_class TNASTYNAGSCREEN    
    SendInput !%ButtonPress%
    ;ListVars
    ;MsgBox, %Button2Press% ;%WTextSlow%
}
