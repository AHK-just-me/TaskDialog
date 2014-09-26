; ======================================================================================================================
; TaskDialog -> msdn.microsoft.com/en-us/library/bb760540(v=vs.85).aspx
; Main:     String to be used for the main instruction (mandatory).
; Extra:    String used for additional text that appears below the main instruction (optional).
;           Default: 0 - no additional text.
; Title:    String to be used for the task dialog title (optional).
;           Default: "" - A_ScriptName.
; Buttons:  Specifies the push buttons displayed in the dialog box (optional).
;           This parameter may be a combination of the integer values defined in TDBTNS or a list of string keys
;           separated by pipe (|), space, comma, or LF (`n).
;           list of the string keys.
;           Default: 0 - OK button
; Icon:     Specifies the icon to display in the task dialog (optional).
;           This parameter can be one of the keys defined in TDICON or a HICON handle to a 32*32 sized icon.
;           Default: 0 - no icon
; Width:    Specifies the width of the task dialog's client area, in pixels.
;           If you pass -1 the width of the task dialog is determined by the width of its content (Extra) area.
;           Default: 0 - the task dialog manager will calculate the ideal width.
; Parent:   HWND of the owner window of the task dialog to be created(optional).
;           If a valid window handle is specified, the task dialog will become modal.
;           Pass -1 to set the task dialog 'AlwaysOnTop'.
;           Default: 0 - no owner window
; Timeout:  Timeout in seconds, which can contain a decimal point.
;           Due to the use of the TDN_TIMER notification the precision will be about 200 ms.
;           Default: 0 - no timeout
; Returns:  An integer value identifying the button pressed by the user:
;           -1 = Timeout, 1 = OK, 2 = CANCEL, 4 = RETRY, 6 = YES, 7 = NO, 8 = CLOSE
;           If the function fails, ErrorLevel will be set and the return value will be 0.
; Remarks:  Depending on the settings of TaskDialogUseMsgBoxOnXP() the function can display a MsgBox instead of
;           the task dialog on Win XP. In that case you should only use icons and in particular button combinations
;           also supported by the MsgBox command:
;              Icon:    1/"WARN", 2/"ERROR", 3/"INFO", "QUESTION"
;              Buttons: 1/"OK", 9/"OK|CANCEL", 6/"YES|NO", 14/"YES|NO|CANCEL", 24/"RETRY|CANCEL"
;           Other icons won't be shown, other button combinations won't show all specified buttons.
; Version:  2014-09-26
; ======================================================================================================================
TaskDialog(Main, Extra := "", Title := "", Buttons := 0, Icon := 0, Width := 0, Parent := 0, TimeOut := 0) {
   Static TDCB      := RegisterCallback("TaskDialogCallback", "Fast")
        , TDCSize   := (4 * 8) + (A_PtrSize * 16)
        , TDBTNS    := {OK: 1, YES: 2, NO: 4, CANCEL: 8, RETRY: 16, CLOSE: 32}
        , TDF       := {HICON_MAIN: 0x0002, ALLOW_CANCEL: 0x0008, CALLBACK_TIMER: 0x0800, SIZE_TO_CONTENT: 0x01000000}
        , TDICON    := {1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9
                      , WARN: 1, ERROR: 2, INFO: 3, SHIELD: 4, BLUE: 5, YELLOW: 6, RED: 7, GREEN: 8, GRAY: 9
                      , QUESTION: 0}
        , HQUESTION := DllCall("User32.dll\LoadIcon", "Ptr", 0, "Ptr", 0x7F02, "UPtr")
        , DBUX      := DllCall("User32.dll\GetDialogBaseUnits", "UInt") & 0xFFFF
        , OffParent := 4
        , OffFlags  := OffParent + (A_PtrSize * 2)
        , OffBtns   := OffFlags + 4
        , OffTitle  := OffBtns + 4
        , OffIcon   := OffTitle + A_PtrSize
        , OffMain   := OffIcon + A_PtrSize
        , OffExtra  := OffMain + A_PtrSize
        , OffCB     := (4 * 7) + (A_PtrSize * 14)
        , OffCBData := OffCB + A_PtrSize
        , OffWidth  := OffCBData + A_PtrSize
   ; -------------------------------------------------------------------------------------------------------------------
   If ((DllCall("Kernel32.dll\GetVersion", "UInt") & 0xFF) < 6) {
      If TaskDialogUseMsgBoxOnXP()
         Return TaskDialogMsgBox(Main, Extra, Title, Buttons, Icon, Parent, Timeout)
      Else {
         MsgBox, 16, %A_ThisFunc%, You need at least Win Vista / Server 2008 to use %A_ThisFunc%().
         ErrorLevel := "You need at least Win Vista / Server 2008 to use " . A_ThisFunc . "()."
         Return 0
      }
   }
   ; -------------------------------------------------------------------------------------------------------------------
   Flags := Width = 0 ? TDF.SIZE_TO_CONTENT : 0
   If (Title = "")
      Title := A_ScriptName
   BTNS := 0
   If Buttons Is Integer
      BTNS := Buttons & 0x3F
   Else
      For Each, Btn In StrSplit(Buttons, ["|", " ", ",", "`n"])
         BTNS |= (B := TDBTNS[Btn]) ? B : 0
   ICO := (I := TDICON[Icon]) ? 0x10000 - I : 0
   If Icon Is Integer
      If ((Icon & 0xFFFF) <> Icon) ; caller presumably passed HICON
         ICO := Icon
   If (Icon = "Question")
      ICO := HQUESTION
   If (ICO > 0xFFFF)
      Flags |= TDF.HICON_MAIN
   AOT := Parent < 0 ? !(Parent := 0) : False ; AlwaysOnTop
   ; -------------------------------------------------------------------------------------------------------------------
   PTitle := A_IsUnicode ? &Title : TaskDialogToUnicode(Title, WTitle)
   PMain  := A_IsUnicode ? &Main : TaskDialogToUnicode(Main, WMain)
   PExtra := Extra = "" ? 0 : A_IsUnicode ? &Extra : TaskDialogToUnicode(Extra, WExtra)
   VarSetCapacity(TDC, TDCSize, 0) ; TASKDIALOGCONFIG structure
   NumPut(TDCSize, TDC, "UInt")
   NumPut(Parent, TDC, OffParent, "Ptr")
   NumPut(BTNS, TDC, OffBtns, "Int")
   NumPut(PTitle, TDC, OffTitle, "Ptr")
   NumPut(ICO, TDC, OffIcon, "Ptr")
   NumPut(PMain, TDC, OffMain, "Ptr")
   NumPut(PExtra, TDC, OffExtra, "Ptr")
   If (AOT) || (TimeOut > 0) {
      If (TimeOut > 0) {
         Flags |= TDF.CALLBACK_TIMER
         TimeOut := Round(Timeout * 1000)
      }
      TD := {AOT: AOT, Timeout: Timeout}
      NumPut(TDCB, TDC, OffCB, "Ptr")
      NumPut(&TD, TDC, OffCBData, "Ptr")
   }
   NumPut(Flags, TDC, OffFlags, "UInt")
   If (Width > 0)
      NumPut(Width * 4 / DBUX, TDC, OffWidth, "UInt")
   If !(RV := DllCall("Comctl32.dll\TaskDialogIndirect", "Ptr", &TDC, "IntP", Result, "Ptr", 0, "Ptr", 0, "UInt"))
      Return TD.TimedOut ? -1 : Result
   ErrorLevel := "The call of TaskDialogIndirect() failed!`nReturn value: " . RV . "`nLast error: " . A_LastError
   Return 0
}
; ======================================================================================================================
; Call this function once passing 1/True if you want a MsgBox to be displayed instead of the task dialog on Win XP.
; ======================================================================================================================
TaskDialogUseMsgBoxOnXP(UseIt := "") {
   Static UseMsgBox := False
   If (UseIt <> "")
      UseMsgBox := !!UseIt
   Return UseMsgBox
}
; ======================================================================================================================
; Internally used functions
; ======================================================================================================================
TaskDialogMsgBox(Main, Extra, Title := "", Buttons := 0, Icon := 0, Parent := 0, TimeOut := 0) {
   Static MBICON := {1: 0x30, 2: 0x10, 3: 0x40, WARN: 0x30, ERROR: 0x10, INFO: 0x40, QUESTION: 0x20}
        , TDBTNS := {OK: 1, YES: 2, NO: 4, CANCEL: 8, RETRY: 16}
   BTNS := 0
   If Buttons Is Integer
      BTNS := Buttons & 0x1F
   Else
      For Each, Btn In StrSplit(Buttons, ["|", " ", ",", "`n"])
         BTNS |= (B := TDBTNS[Btn]) ? B : 0
   Options := 0
   Options |= (I := MBICON[Icon]) ? I : 0
   Options |= Parent = -1 ? 262144 : Parent > 0 ? 8192 : 0
   If ((BTNS & 14) = 14)
      Options |= 0x03 ; Yes/No/Cancel
   Else If ((BTNS & 6) = 6)
      Options |= 0x04 ; Yes/No
   Else If ((BTNS & 24) = 24)
      Options |= 0x05 ; Retry/Cancel
   Else If ((BTNS & 9) = 9)
      Options |= 0x01 ; OK/Cancel
   Main .= Extra <> "" ? "`n`n" . Extra : ""
   MsgBox, % Options, %Title%, %Main%, %TimeOut%
   IfMsgBox, OK
      Return 1
   IfMsgBox, Cancel
      Return 2
   IfMsgBox, Retry
      Return 4
   IfMsgBox, Yes
      Return 6
   IfMsgBox, No
      Return 7
   IfMsgBox, TimeOut
      Return -1
   Return 0
}
; ======================================================================================================================
TaskDialogToUnicode(String, ByRef Var) {
   VarSetCapacity(Var, StrPut(String, "UTF-16") * 2, 0)
   StrPut(String, &Var, "UTF-16")
   Return &Var
}
; ======================================================================================================================
TaskDialogCallback(H, N, W, L, D) {
   Static TDM_CLICK_BUTTON := 0x0466
        , TDN_CREATED := 0
        , TDN_TIMER   := 4
   TD := Object(D)
   If (N = TDN_TIMER) && (W > TD.Timeout) {
      TD.TimedOut := True
      PostMessage, %TDM_CLICK_BUTTON%, 2, 0, , ahk_id %H% ; IDCANCEL = 2
   }
   Else If (N = TDN_CREATED) && TD.AOT {
      DHW := A_DetectHiddenWindows
      DetectHiddenWindows, On
      WinSet, AlwaysOnTop, On, ahk_id %H%
      DetectHiddenWindows, %DHW%
   }
   Return 0
}