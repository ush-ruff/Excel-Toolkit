#Requires AutoHotkey v2.0
#SingleInstance Force

; ----------------------------------------
; Tray configuration
; ----------------------------------------
A_TrayMenu.Delete()
A_TrayMenu.Add("Open", (*) => Edit())
A_TrayMenu.Add("Reload Script", (*) => Reload())
A_TrayMenu.Add("Pause Hotkeys", (*) => Suspend())
A_TrayMenu.Add("Exit", (*) => ExitApp())

; ----------------------------------------
; Shortcuts (active only when Excel is focused)
; ----------------------------------------
#HotIf WinActive("ahk_exe EXCEL.EXE")

; --- Merge & Unmerge ---
; Merge across selection
; Key: Ctrl + Alt + M
^!m::MergeAcross()

; Merge
; Key: Ctrl + Alt + Shift + M
^!+m::Merge()

; Unmerge cells
; Key: Ctrl + Shift + M
^+u::UnmergeCells()

; --- Fill Colors ---
; Clear fill
; Key: Ctrl + Alt + F
^!f::ClearFill()

; Blue fill
; Key: Ctrl + Alt + G
^!g::SetFillColor("blue")

; Red fill
; Key: Ctrl + Alt + Y
^!y::SetFillColor("red")

; White fill
; Key: Ctrl + Alt + E
^!e::SetFillColor("white")

; Light blue fill
; Key: Ctrl + Alt + I
^!i::SetFillColor("light-blue")

; --- Font Colors ---
; Automatic font color
; Key: Ctrl + Alt + W
^!w::SetFontColorAuto()

; Blue font
; Key: Ctrl + Alt + L
^!l::SetFontColor("blue")

; Red font
; Key: Ctrl + Alt + R
^!r::SetFontColor("red")

; White font
; Key: Ctrl + Alt + Q
^!q::SetFontColor("white")

; --- Text Toggles ---
; Wrap text toggle
; Key: Ctrl + Alt + Z
^!z::ToggleWrap()

; --- Alignment ---
; Left align
; Key: Ctrl + Alt + Left
^!Left::SetHAlign("left")

; Right align
; Key: Ctrl + Alt + Right
^!Right::SetHAlign("right")

; Center align
; Key: Ctrl + Alt + C
^!c::SetHAlign("center")

; Top align
; Key: Ctrl + Alt + Up
^!Up::SetVAlign("top")

; Down align
; Key: Ctrl + Alt + Down
^!Down::SetVAlign("down")

; Middle align
; Key: Ctrl + Shift + C
^+c::SetVAlign("middle")

; --- Font Size ---
; Increase font size by 1
; Key: Ctrl + Alt + NumpadAdd
^!NumpadAdd::IncreaseFont()

; Decrease font size by 1
; Key: Ctrl + Alt + NumpadSub
^!NumpadSub::DecreaseFont()

; --- Number Formats ---
; Key: Ctrl + Alt + Numpad1
^!Numpad1::SetNumberFormat("Accounting")

; General format
; Key: Ctrl + Alt + Numpad2
^!Numpad2::SetNumberFormat("General")

; Percentage format
; Key: Ctrl + Alt + Numpad3
^!Numpad3::SetNumberFormat("Percentage")

; Date format
; Key: Ctrl + Alt + 4
; ^!4::SetNumberFormat("yyyy-mm-dd")

; --- Borders ---
; Remove borders
; Key: Ctrl + Alt + Shift + B
^!+b::SetBorders("Remove")

; Outside borders
; Key: Ctrl + Alt + B
^!b::SetBorders("Outside")

; Key: Ctrl + Shift + B
^+b::SetBorders("All")

; --- Freeze Panes ---
; Toggle freeze panes at active cell
; Key: Ctrl + Alt + F9
^!f9::ToggleFreezePanes()

; --- Quick Actions ---
; Clean format (remove fill, bold, wrap, reset font)
; Key: Ctrl + Alt + K
^!k::QuickClean()

; AutoFit columns and rows for selection
; Key: Ctrl + Alt + A
^!a::AutoFitSelection()

; Delete row
; Key: Ctrl + Shift + K
^+k::DeleteRow()

#HotIf

; ============================================================
; Helper Functions
; ============================================================
; Exits edit mode then sends keystrokes (so they act on the sheet, not a cell)
SendKeys(keys) {
  Send "{Esc}" keys
}

; ============================================================
; Merge & Unmerge
; ============================================================
MergeAcross() {
  SendKeys("!hma")
}

Merge() {
  SendKeys("!hmm")
}

UnmergeCells() {
  SendKeys("!hmu")
}

; ============================================================
; Fill & Font (COM — cosmetic, no undo)
; ============================================================
ClearFill() {
  SendKeys("!hhn")
}

SetFillColor(val) {
  switch val {
    case "blue": SendKeys("!hhm{Right}{Tab}{Tab}{Tab}{Tab}0{Tab}0{Tab}255{Enter}")
    case "red": SendKeys("!hhm{Right}{Tab}{Tab}{Tab}{Tab}255{Tab}0{Tab}0{Enter}")
    case "white": SendKeys("!hhm{Right}{Tab}{Tab}{Tab}{Tab}255{Tab}255{Tab}255{Enter}")
    case "light-blue": SendKeys("!hhm{Right}{Tab}{Tab}{Tab}{Tab}135{Tab}206{Tab}250{Enter}")
  }
}

SetFontColorAuto() {
  SendKeys("!hfc{Space}")
}

SetFontColor(val) {
  switch val {
    case "blue": SendKeys("!hfcm{Right}{Tab}{Tab}{Tab}{Tab}0{Tab}0{Tab}255{Enter}")
    case "red": SendKeys("!hfcm{Right}{Tab}{Tab}{Tab}{Tab}255{Tab}0{Tab}0{Enter}")
    case "white": SendKeys("!hfcm{Right}{Tab}{Tab}{Tab}{Tab}255{Tab}255{Tab}255{Enter}")
  }
}

; ============================================================
; Text Toggles
; ============================================================
ToggleWrap() {
  SendKeys("!hw")
}

; ============================================================
; Alignment
; ============================================================
SetHAlign(val) {
  switch val {
    case "left": SendKeys("!hal")
    case "right": SendKeys("!har")
    case "center": SendKeys("!hac")
  }
}

SetVAlign(val) {
  switch val {
    case "top": SendKeys("!hat")
    case "down": SendKeys("!hab")
    case "middle": SendKeys("!ham")
  }
}

; ============================================================
; Font Size
; ============================================================
IncreaseFont() {
  SendKeys("!hfg")
}

DecreaseFont() {
  SendKeys("!hfk")
}

; ============================================================
; Number Formats
; ============================================================
SetNumberFormat(val) {
  switch val {
    case "Accounting": SendKeys("!hk!h9!h9")
    case "General": SendKeys("!hnGeneral{Enter}")
    case "Percentage": SendKeys("!hp")
  }
}

; ============================================================
; Borders
; ============================================================
SetBorders(val) {
  switch val {
    case "Remove": SendKeys("!hbn")
    case "Outside": SendKeys("!hbs")
    case "All": SendKeys("!hba")
  }
}

; ============================================================
; Freeze Panes
; ============================================================
ToggleFreezePanes() {
  SendKeys("!wff")
}

; ============================================================
; Quick Actions
; ============================================================
QuickClean() {
  SendKeys("!hef")
}

AutoFitSelection() {
  SendKeys("!hoi")
  Sleep 30
  SendKeys("!hoa")
}

DeleteRow() {
  SendKeys("!hdr")
}
