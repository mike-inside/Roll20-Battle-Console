
MsgBoxEx(Text, Title := "", Buttons := "", Icon := "", ByRef CheckText := "", Styles := "", Owner := "", Timeout := "", FontOptions := "", FontName := "", BGColor := "", Callback := "") {
    Static hWnd, y2, p, px, pw, c, c2, cw, cy, ch, f, o, gL, hBtn, lb, DHW, w, ww, Off, k, v, RetVal
    Static Sound := {2: "*48", 4: "*16", 5: "*64"}

    Gui New, hWndhWnd LabelMsgBoxEx -0xA0000
    Gui % (Owner) ? "+Owner" . Owner : ""
    Gui Font
    Gui Font, % (FontOptions) ? FontOptions : "s9", % (FontName) ? FontName : "Segoe UI"
    Gui Color, % (BGColor) ? BGColor : "White"
    Gui Margin, 10, 12

    If (IsObject(Icon)) {
        Gui Add, Picture, % "x20 y24 w32 h32 Icon" . Icon[1], % (Icon[2] != "") ? Icon[2] : "shell32.dll"
    } Else If (Icon + 0) {
        Gui Add, Picture, x20 y24 Icon%Icon% w32 h32, user32.dll
        SoundPlay % Sound[Icon]
    }

    Gui Add, Link, % "x" . (Icon ? 65 : 20) . " y" . (InStr(Text, "`n") ? 24 : 32) . " vc", %Text%
    GuicontrolGet c, Pos
	GuiControl Move, c, % "w" . (cw + 30)
    if (cw > 500) {
		GuiControl Hide, c
		Gui Add, Link, % "x" . (Icon ? 65 : 20) . " y" . (InStr(Text, "`n") ? 24 : 32) . " vc2 w500", %Text%
		GuicontrolGet c, Pos, c2
	}
    y2 := (cy + ch < 52) ? 90 : cy + ch + 34

    Gui Add, Text, vf -Background ; Footer

    Gui Font
    Gui Font, s9, Segoe UI
    px := 42
    If (CheckText != "") {
        CheckText := StrReplace(CheckText, "*",, ErrorLevel)
        Gui Add, CheckBox, vCheckText x12 y%y2% h26 -Wrap -Background AltSubmit Checked%ErrorLevel%, %CheckText%
        GuicontrolGet p, Pos, CheckText
        px := px + pw + 10
    }

    o := {}
    Loop Parse, Buttons, |, *
    {
        gL := (Callback != "" && InStr(A_LoopField, "...")) ? Callback : "MsgBoxExBUTTON"
		w := InStr(Title, "Special Abilities") ? "" : "w90" 
        Gui Add, Button, hWndhBtn g%gL% x%px% %w% y%y2% h26 -Wrap, %A_Loopfield%
		GuicontrolGet c, Pos, %A_Loopfield%
        lb := hBtn
        o[hBtn] := px
        ;px += 98
		px := InStr(Title, "Special Abilities") ? px + cw + 8 : px + 98
    }
    GuiControl +Default, % (RegExMatch(Buttons, "([^\*\|]*)\*", Match)) ? Match1 : StrSplit(Buttons, "|")[1]

    Gui Show, Autosize Center Hide, %Title%
    DHW := A_DetectHiddenWindows
    DetectHiddenWindows On
    WinGetPos,,, ww,, ahk_id %hWnd%
    GuiControlGet p, Pos, %lb% ; Last button
    Off := ww - (((px + pw + 14) * A_ScreenDPI) // 96)
    For k, v in o {
        GuiControl Move, %k%, % "x" . (v + Off)
    }
    Guicontrol MoveDraw, f, % "x-1 y" . (y2 - 10) . " w" . ww . " h" . 48

    Gui Show
    Gui +SysMenu %Styles%
    DetectHiddenWindows %DHW%

    If (Timeout) {
        SetTimer MsgBoxExTIMEOUT, % Round(Timeout) * 1000
    }

    If (Owner) {
        WinSet Disable,, ahk_id %Owner%
    }

    GuiControl Focus, f
    Gui Font
    WinWaitClose ahk_id %hWnd%
    Return RetVal

    MsgBoxExESCAPE:
    MsgBoxExCLOSE:
    MsgBoxExTIMEOUT:
    MsgBoxExBUTTON:
        SetTimer MsgBoxExTIMEOUT, Delete

        If (A_ThisLabel == "MsgBoxExBUTTON") {
            RetVal := StrReplace(A_GuiControl, "&")
        } Else {
            RetVal := (A_ThisLabel == "MsgBoxExTIMEOUT") ? "Timeout" : "Cancel"
        }

        If (Owner) {
            WinSet Enable,, ahk_id %Owner%
        }

        Gui Submit
        Gui %hWnd%: Destroy
    Return
}

SoftModalMessageBox(Text, Title, Buttons, DefBtn := 1, Options := 0x1, IconRes := "", IconID := 1, Timeout := -1, Owner := 0, Callback := "") {

    If (IconRes != "") {
        hModule := DllCall("GetModuleHandle", "Str", IconRes, "Ptr")
        LoadLib := !hModule
            && hModule := DllCall("kernel32.dll\LoadLibraryEx", "Str", IconRes, "UInt", 0, "UInt", 0x2, "Ptr")
        Options |= 0x80 ; MB_USERICON
    } Else {
        hModule := 0
        LoadLib := False
    }

    cButtons := Buttons.Length()
    VarSetCapacity(ButtonIDs, cButtons * A_PtrSize, 0)
    VarSetCapacity(ButtonText, cButtons * A_PtrSize, 0)
    Loop %cButtons% {
        NumPut(Buttons[A_Index][1], ButtonIDs, 4 * (A_Index - 1), "UInt")
        NumPut(&(b%A_Index% := Buttons[A_Index][2]), ButtonText, A_PtrSize * (A_Index - 1), "Ptr")
    }

    If (Callback != "") {
        Callback := RegisterCallback(Callback, "F")
    }

    x64 := A_PtrSize == 8
    Offsets := (A_Is64BitOS) ? (x64 ? [96, 104, 112, 116, 120, 124] : [52, 56, 60, 64, 68, 72]) : [48, 52, 56, 60, 64, 68]

    ; MSGBOXPARAMS and MSGBOXDATA structures
    NumPut(VarSetCapacity(MBCONFIG, (x64) ? 136 : 76, 0), MBCONFIG, 0, "UInt")
    NumPut(Owner,    MBCONFIG, 1 * A_PtrSize, "Ptr")  ; Owner window
    NumPut(hModule,  MBCONFIG, 2 * A_PtrSize, "Ptr")  ; Icon resource
    NumPut(&Text,    MBCONFIG, 3 * A_PtrSize, "Ptr")  ; Message
    NumPut(&Title,   MBCONFIG, 4 * A_PtrSize, "Ptr")  ; Window title
    NumPut(Options,  MBCONFIG, 5 * A_PtrSize, "UInt") ; Options
    NumPut(IconID,   MBCONFIG, 6 * A_PtrSize, "Ptr")  ; Icon resource ID
    NumPut(Callback, MBCONFIG, 8 * A_PtrSize, "Ptr")  ; Callback
    NumPut(&ButtonIDs,  MBCONFIG, Offsets[1], "Ptr")  ; Button IDs
    NumPut(&ButtonText, MBCONFIG, Offsets[2], "Ptr")  ; Button texts
    NumPut(cButtons,    MBCONFIG, Offsets[3], "UInt") ; Number of buttons
    NumPut(DefBtn - 1,  MBCONFIG, Offsets[4], "UInt") ; Default button
    NumPut(1,           MBCONFIG, Offsets[5], "UInt") ; Allow cancellation
    NumPut(Timeout,     MBCONFIG, Offsets[6], "Int")  ; Timeout (ms)

    ProcAddr := DllCall("GetProcAddress", "Ptr", DllCall("GetModuleHandle", "Str", "User32.dll", "Ptr"), "AStr", "SoftModalMessageBox", "Ptr")
    Ret := DllCall(ProcAddr, "Ptr", &MBCONFIG)

    If (LoadLib) {
        DllCall("FreeLibrary", "Ptr", hModule)
    }

    If (Callback != "") {
        DllCall("GlobalFree", "Ptr", Callback)
    }

    Return Ret
}