;=======================================================================================
; GdOcrTool.ahk is an AutoHotkey (v1.1) script that enables users to easily capture a  ;
; single word or an area and send to GoldenDict. Capture2Text is required.             ;
; Written by Johnny Van, 2021/11/13                                                    ;
; Updated 2021/11/14, added support for MDict, Eudic.                                  ;
; Updated 2021/11/19, added visual feedback for single word capture.                   ;
; Updated 2021/12/5, fixed bug that caused mouse frozen, improved success rate for     ;
; single word captrue, added configuration file, added cross-shaped cursor, added      ;
; zoom-in feature.                                                                     ;
;=======================================================================================
; Auto-execution section.

#NoEnv
#SingleInstance, force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
CoordMode, ToolTip, Screen
CoordMode, Mouse, Screen
DetectHiddenWindows, On
;FileEncoding, UTF-8

Global GdOcrToolConfigFileName := A_ScriptDir . "\GdOcrTool.ini"
Global Capture2TextFileName := ""
Global Capture2TextConfigFileName := ""
Global GoldenDictFileName := ""
Global MdictFileName := ""
Global EudicFileName := ""
Global DefaultCapture2TextFileName := A_ScriptDir . "\Capture2Text.exe"
Global DefaultGoldenDictFileName := SubStr(A_ScriptDir, 1, -StrLen("Capture2Text")) . "GoldenDict\GoldenDict.exe"
Global DefaultMdictFileName := "C:\Program Files\MDictPC\MDict.exe"
Global DefaultEudicFileName := "C:\Program Files (x86)\eudic\eudic.exe"
Global Capture2TextAhkExe := "ahk_exe Capture2Text.exe"
Global GoldenDictAhkExe := "ahk_exe GoldenDict.exe"
Global MdictAhkExe := "ahk_exe MDict.exe"
Global EudicAhkExe := "ahk_exe eudic.exe"
Global DictApp := "GoldenDict"  ; Dictionary app: "GoldenDict", "MDict", "Eudic"
Global DictAppGroup := {}
Global CaptureMode := "NoCapture"  ; Capture mode: "NoCapture", "SingleWordCapture", "BoxCapture"
Global SingleWordCaptureCount := 0
Global LineCaptured := ""
Global ForwardLineCaptured := ""
Global SingleWordCaptureTimeout := 1000  ; Timeout in millisecond. Abort capture if the timeout has expired.
Global BoxCaptureTimeout := 10000
Global CaptureTimeoutArray := Object("NoCapture", 0, "SingleWordCapture", SingleWordCaptureTimeout, "BoxCapture", BoxCaptureTimeout)
Global StartTime, EndTime
Global IsDrawBox := False
Global EnableZoom := False
Global EnableDebugInfo := False
Global EnableTimeoutPrompt := False
Global OmitChars := " ``~!@#$%^&*()_=+[{]}\|;:"",<.>/?·！@￥…（）—【】、；：‘“’”，《。》？"
Global DelimiterArray := StrSplit(OmitChars)
Global GuiOverlayHwnd := ""

Main()

; End of auto-execution section.
;=======================================================================================
; Functions

Main() {
    Menu, Tray, Icon, shell32.dll, 172
    LoadConfig()
    SaveConfig()
    CreateGuiOverlay()
    OnClipboardChange("ClipboardChange")
    ;OnExit("CloseApps")
    Return
}

LoadConfig() {
    DictAppObj := {}
    DictAppObj.FileName := GoldenDictFileName
    DictAppObj.DefaultFileName := DefaultGoldenDictFileName
    DictAppObj.AhkExe := GoldenDictAhkExe
    DictAppGroup["GoldenDict"] := DictAppObj

    DictAppObj := {}
    DictAppObj.FileName := MdictFileName
    DictAppObj.DefaultFileName := DefaultMdictFileName
    DictAppObj.AhkExe := MdictAhkExe
    DictAppGroup["MDict"] := DictAppObj

    DictAppObj := {}
    DictAppObj.FileName := EudicFileName
    DictAppObj.DefaultFileName := DefaultEudicFileName
    DictAppObj.AhkExe := EudicAhkExe
    DictAppGroup["Eudic"] := DictAppObj

    If FileExist(GdOcrToolConfigFileName) {
        IniRead, DictApp, %GdOcrToolConfigFileName%, User Config, DictApp
        IniRead, EnableZoom, %GdOcrToolConfigFileName%, User Config, EnableZoom
        IniRead, EnableDebugInfo, %GdOcrToolConfigFileName%, User Config, EnableDebugInfo
        IniRead, EnableTimeoutPrompt, %GdOcrToolConfigFileName%, User Config, EnableTimeoutPrompt
        IniRead, SingleWordCaptureTimeout, %GdOcrToolConfigFileName%, User Config, SingleWordCaptureTimeout
        IniRead, BoxCaptureTimeout, %GdOcrToolConfigFileName%, User Config, BoxCaptureTimeout
        IniRead, Capture2TextFileName, %GdOcrToolConfigFileName%, User Config, Capture2TextFileName
        IniRead, GoldenDictFileName, %GdOcrToolConfigFileName%, User Config, GoldenDictFileName
        IniRead, MdictFileName, %GdOcrToolConfigFileName%, User Config, MdictFileName
        IniRead, EudicFileName, %GdOcrToolConfigFileName%, User Config, EudicFileName

        DictAppGroup["GoldenDict"].FileName := GoldenDictFileName
        DictAppGroup["MDict"].FileName := MdictFileName
        DictAppGroup["Eudic"].FileName := EudicFileName
        CaptureTimeoutArray["SingleWordCapture"] := SingleWordCaptureTimeout
        CaptureTimeoutArray["BoxCapture"] := BoxCaptureTimeout
    } Else {
        Capture2TextFileName := DefaultCapture2TextFileName
        GoldenDictFileName := DefaultGoldenDictFileName
        MdictFileName := DefaultMdictFileName
        EudicFileName := DefaultEudicFileName
    }

    If !FileExist(Capture2TextFileName) {
        Capture2TextFileName := DefaultCapture2TextFileName
        If !FileExist(Capture2TextFileName) {
            MsgBox, 48, Warning, Capture2Text.exe not found!`nSpecify its location.
            FileSelectFile, Capture2TextFileName, 3, %A_ScriptDir%, Select Capture2Text.exe, Executable (*.exe)
        }
    }
    If WinExist(Capture2TextAhkExe) {
        WinGet, Capture2TextPid, PID, ahk_exe Capture2Text.exe
        Process, Close, %Capture2TextPid%  ; Kill Capture2Text process to allow --portable option.
    }
    Capture2TextConfigFileName := SubStr(Capture2TextFileName, 1, -StrLen("Capture2Text.exe")) . "Capture2Text\Capture2Text.ini"
    RunCapture2Text()

    DictAppFileName := DictAppGroup[DictApp].FileName
    If !FileExist(DictAppFileName) {
        DictAppFileName := DictAppGroup[DictApp].DefaultFileName
        If !FileExist(DictAppFileName) {
            MsgBox, 48, Warning, %DictApp% not found!`nSpecify its location.
            FileSelectFile, DictAppFileName, 3, %A_ScriptDir%, Select %DictApp%, Executable (*.exe)
        }
        DictAppGroup[DictApp].FileName := DictAppFileName
    }
    If !WinExist(DictAppGroup[DictApp].AhkExe) {
        Run % DictAppFileName
    }

    Return
}

SaveConfig() {
    GoldenDictFileName := DictAppGroup["GoldenDict"].FileName
    MdictFileName := DictAppGroup["MDict"].FileName
    EudicFileName := DictAppGroup["Eudic"].FileName

    IniWrite, %DictApp%, %GdOcrToolConfigFileName%, User Config, DictApp
    IniWrite, %EnableZoom%, %GdOcrToolConfigFileName%, User Config, EnableZoom
    IniWrite, %EnableDebugInfo%, %GdOcrToolConfigFileName%, User Config, EnableDebugInfo
    IniWrite, %EnableTimeoutPrompt%, %GdOcrToolConfigFileName%, User Config, EnableTimeoutPrompt
    IniWrite, %SingleWordCaptureTimeout%, %GdOcrToolConfigFileName%, User Config, SingleWordCaptureTimeout
    IniWrite, %BoxCaptureTimeout%, %GdOcrToolConfigFileName%, User Config, BoxCaptureTimeout
    IniWrite, %Capture2TextFileName%, %GdOcrToolConfigFileName%, User Config, Capture2TextFileName
    IniWrite, %GoldenDictFileName%, %GdOcrToolConfigFileName%, User Config, GoldenDictFileName
    IniWrite, %MdictFileName%, %GdOcrToolConfigFileName%, User Config, MdictFileName
    IniWrite, %EudicFileName%, %GdOcrToolConfigFileName%, User Config, EudicFileName

    If ErrorLevel {
        MsgBox, 16, Error, Failed to write file.`nYou might need to run GdOcrTool as administrator.
        ExitApp
    }
    Return
}

; Create a transparent window to prevent mouse cursor interacting with background window.
CreateGuiOverlay() {
    Gui, GuiOverlay: New, +AlwaysOnTop +Disabled -Caption HwndGuiOverlayHwnd
    WinSet, Transparent, 1, ahk_id %GuiOverlayHwnd%
    Return
}

ClipboardChange(Type) {
; 0 = Clipboard is now empty.
; 1 = Clipboard contains something that can be expressed as text (this includes files copied from an Explorer window).
; 2 = Clipboard contains something entirely non-text such as a picture.
    If (Type != 1) {
        Return
    }

    Switch CaptureMode {
        Case "NoCapture":
            Return
        Case "SingleWordCapture":
            SingleWordCaptureHandler()
        Case "BoxCapture":
            BoxCaptureHandler()
    }
    Return
}

SingleWordCaptureHandler() {
    SingleWordCaptureCount += 1
    Switch SingleWordCaptureCount {
        Case 1:
            ; Clipboard changed by text line capture.
            LineCaptured := Clipboard
            ForwardLineCaptureRequest()
        Case 2:
            ; Clipboard changed by forward text line capture.
            ResetCapture()
            ForwardLineCaptured := Clipboard
            SearchTerm := ExtractSingleWord(LineCaptured, LTrim(ForwardLineCaptured, OmitChars))
            If (SearchTerm != "") {
                SendToDictApp(SearchTerm)
                TemporaryToolTip(SearchTerm, -1000)
            } Else {
                TemporaryToolTip("Failed to extract.", -1000)
            }

            MouseGetPos, MousePosX, MousePosY
            EndTime := A_TickCount
            ElapsedTime := EndTime - StartTime
            If EnableDebugInfo {
                DebugString := "Line captured: [" . LineCaptured . "]`nForward line captured: [" . ForwardLineCaptured . "]`nWord extracted: [" . SearchTerm . "]`nElapsed time: [" . ElapsedTime . "ms]`nPosition: [" . MousePosX . ", " . MousePosY . "]"
                MsgBox, 64, Debug Info, %DebugString%
            }
        Default:
            ResetCapture()
    }
    Return
}

BoxCaptureHandler() {
    ResetCapture()
    SearchTerm := Clipboard
    If ((SearchTerm != "") And (SearchTerm != "<Error>")) {
        SendToDictApp(SearchTerm)
    }

    EndTime := A_TickCount
    ElapsedTime := EndTime - StartTime
    If EnableDebugInfo {
        DebugString := "Text captured: [" . SearchTerm . "]`nElapsed time: [" . ElapsedTime . "ms]"
        MsgBox 64, Debug Info, %DebugString%
    }
    Return
}

ExtractSingleWord(TextLine, ForwardTextLine) {
    SingleWord := ExtractMethodOne(TextLine, ForwardTextLine)
    If (SingleWord == "") {
        SingleWord := ExtractMethodTwo(TextLine, ForwardTextLine)
    }
    Return SingleWord
}

ExtractMethodOne(TextLine, ForwardTextLine) {
    SingleWord := ""

    ForwardTextLinePos := InStr(TextLine, ForwardTextLine)
    If (ForwardTextLinePos != 0) {
        FrontString := SubStr(TextLine, 1, ForwardTextLinePos-1)
        ArrayTemp := StrSplit(FrontString, DelimiterArray)
        SingleWordFront := ArrayTemp[ArrayTemp.Length()]
        SingleWordEnd := StrSplit(ForwardTextLine, DelimiterArray)[1]
        SingleWord := SingleWordFront . SingleWordEnd
    }

    If ((SingleWord == "") And (StrSplit(ForwardTextLine, DelimiterArray).Length() > 1)) {
        SingleWord := ExtractMethodOne(TextLine, SubStr(ForwardTextLine, 1, StrLen(ForwardTextLine)-1))
    }

    Return SingleWord
}

ExtractMethodTwo(TextLine, ForwardTextLine) {
    SingleWord := ""

    SingleWordCandidateArray := StrSplit(TextLine, DelimiterArray)
    ForwardPartialWord := StrSplit(ForwardTextLine, DelimiterArray)[1]

    MaxCharMatchIndex := 0
    CharMatchArray := []
    Loop % SingleWordCandidateArray.Length() {
        CharMatch := 0
        SingleWordCandidate := SingleWordCandidateArray[A_Index]

        If (StrLen(ForwardPartialWord) <= StrLen(SingleWordCandidate)) {
            Loop % StrLen(ForwardPartialWord) {
                If (SubStr(ForwardPartialWord, -(A_Index-1), 1) == SubStr(SingleWordCandidate, -(A_Index-1), 1)) {
                    CharMatch += 1
                }
            }
        }
        CharMatchArray[A_Index] := CharMatch
        If (CharMatch > CharMatchArray[MaxCharMatchIndex]) {
            MaxCharMatchIndex := A_Index
        }
    }

    If (MaxCharMatchIndex != 0) {
        SingleWord := SingleWordCandidateArray[MaxCharMatchIndex]
    }
    
    Return SingleWord
}

RunCapture2Text() {
    If FileExist(Capture2TextConfigFileName) {
        Run, %Capture2TextFileName% --portable
    } Else {
        Run % Capture2TextFileName
    }
    Return
}

SendToDictApp(SearchTerm) {
    Switch DictApp {
        Case "GoldenDict":
            SendToGoldenDict(SearchTerm)
        Case "MDict":
            SendToMDict(SearchTerm)
        Case "Eudic":
            SendToEudic(SearchTerm)
    }
    Return
}

SendToGoldenDict(SearchTerm) {
    SearchTermCli := """" . StrReplace(SearchTerm, """", """""""") . """"  ; Escape double quotes
    Run, %GoldenDictFileName% %SearchTermCli%
    Return
}

SendToMDict(SearchTerm) {
    Clipboard := SearchTerm
    Run, %MdictFileName%
    WinWait, %MdictAhkExe%, , 0.2
    If WinActive(MdictAhkExe) {
        Send, ^v
        Sleep, 50
        Send, {Enter}
    }
    Return
}

SendToEudic(SearchTerm) {
    Clipboard := SearchTerm
    Run, %EudicFileName%
    WinWait, %EudicAhkExe%, , 0.2
    If WinActive(EudicAhkExe) {
        Send, ^v
        Sleep, 50
        Send, {Enter}
    }
    Return
}

ResetCapture() {
    CaptureMode := "NoCapture"
    SingleWordCaptureCount := 0
    SetTimer, AbortOverdueCapture, Off
    SetTimer, TurnOffToolTip, Off
    Return
}

AbortOverdueCapture() {
    EndTime := A_TickCount
    ElapsedTime := EndTime - StartTime
    If (ElapsedTime > CaptureTimeoutArray[CaptureMode]) {
        ResetCapture()
        If EnableTimeoutPrompt {
            TemporaryToolTip("Timeout has expired.", -1000)
        }
    }
    Return
}

TemporaryToolTip(ToolTipText, Period) {
    ToolTip % ToolTipText
    SetTimer, TurnOffToolTip, %Period%
    Return
}

TurnOffToolTip() {
    ToolTip
    Return
}

StartZoom() {
    Run, Magnify.exe
    WinWait, ahk_exe Magnify.exe, , 2
    Return
}

EndZoom() {
    PostMessage, 0x0112, 0xF060, , , ahk_exe Magnify.exe
    Return
}

; Replace arrow-shaped cursor (32512) with cross-shaped cursor (32515).
SetCrossCursor() {
    hCursor := DllCall("LoadImage", "Uint", 0, "Uint", 32515, "Uint", 2, "Uint", 0, "Uint", 0, "Uint", 0x8000)
    DllCall("SetSystemCursor", "Uint", DllCall("CopyImage", "Uint", hCursor, "Uint", 2, "Int", 0, "Int", 0, "Uint", 0), "Uint", 32512)
    Return
}

ReloadSystemCursor() {
    DllCall("SystemParametersInfo", "Uint", 0x0057, "Uint", 0, "Uint", 0, "Uint", 0)
    Return
}

; Call Capture2Text by sending hotkeys.
LineCaptureRequest() {
    Send, ^+#e  ; crtl + shift + win + e
    Return
}

ForwardLineCaptureRequest() {
    Send, ^+#w  ; crtl + shift + win + w
    Return
}

BoxCaptureRequest() {
    Send, ^+#q  ; crtl + shift + win + q
    Return
}

CloseApps() {
    DictAppAhkExe := DictAppGroup[DictApp].AhkExe
    WinGet, DictAppPid, PID, %DictAppAhkExe%
    WinGet, Capture2TextPid, PID, ahk_exe Capture2Text.exe

    Process, Close, %DictAppPid%
    Process, Close, %Capture2TextPid%
    Return
}

;=======================================================================================
; GdOcrTool hotkeys.

^RButton::  ; Capture a single word by pressing ctrl + right click.
SingleWordCapture() {
    If !WinExist(Capture2TextAhkExe) {
        MsgBox, 48, Warning, Capture2Text is not running!
        RunCapture2Text()
        Return
    }
    ResetCapture()
    StartTime := A_TickCount
    CaptureMode := "SingleWordCapture"
    LineCaptureRequest()
    SetTimer, AbortOverdueCapture, 100
    Return
}

^`::  ; Start box capture by pressing ctrl + `
BoxCapture() {
    If !WinExist(Capture2TextAhkExe) {
        MsgBox, 48, Warning, Capture2Text is not running!
        RunCapture2Text()
        Return
    }
    ResetCapture()
    IsDrawBox := True
    Gui, GuiOverlay: Show, Maximize
    If EnableZoom {
        StartZoom()
    }
    SetCrossCursor()
    Return
}

; Context-sensitive hotkeys.
#If IsDrawBox

LButton::
StartDrawBox() {
    TurnOffToolTip()
    BoxCaptureRequest()
    Return
}

LButton Up::
EndDrawBox() {
    Send, {LButton Down}
    IsDrawBox := False
    CaptureMode := "BoxCapture"
    StartTime := A_TickCount  ; Start count down after box is drawn.
    SetTimer, AbortOverdueCapture, 500
    If EnableZoom {
        EndZoom()
    }
    ReloadSystemCursor()
    Gui, GuiOverlay: Hide
    Return
}

Esc::
ForceAbortBoxCapture() {
    IsDrawBox := False
    Gui, GuiOverlay: Hide
    If EnableZoom {
        EndZoom()
    }
    ReloadSystemCursor()
    ResetCapture()
    Return
}

#If

;=======================================================================================
