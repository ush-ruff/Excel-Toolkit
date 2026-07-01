; Excel Toolkit
; Press Win + ` to toggle the hotstrings off
; Script v1.2.1

#Requires AutoHotkey v2.0
#SingleInstance Force
SetWorkingDir A_ScriptDir

; ----------------------------------------
; Tray configuration
; ----------------------------------------
A_TrayMenu.Delete()
A_TrayMenu.Add("Open", (*) => Edit())
A_TrayMenu.Add("Toggle Hotstrings", (*) => ToggleHotstrings())
A_TrayMenu.Add()
A_TrayMenu.Add("Reload Script", (*) => Reload())
A_TrayMenu.Add("Pause Hotkeys", (*) => Suspend())
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Open"
A_IconTip := "Excel Toolkit - Hotstrings ENABLED"

; ----------------------------------------
; Global Variables
; ----------------------------------------
hotstringsEnabled := true
Snippets := LoadSnippets(A_ScriptDir "\snippets.ini", "Excel")

; Format: "name", "R, G, B"  — each value 0-255 (comma + space required)
colors := Map(
	"blue", "0, 0, 255",
	"red", "192, 0, 0",
	"white", "255, 255, 255",
	"rose", "255, 204, 204",
	"yellow", "255, 255, 0",
	"gray", "217, 217, 217",
	"light-blue", "221, 235, 247"
)

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
^+m::UnmergeCells()

; --- Fill Colors ---
; Clear fill
; Key: Ctrl + Alt + Numpad0
^!Numpad0::ClearFill()

; Light blue fill
; Key: Ctrl + Alt + Numpad1
^!Numpad1::SetFillColor("light-blue")

; Rose fill
; Key: Ctrl + Alt + Numpad2
^!Numpad2::SetFillColor("rose")

; Gray fill
; Key: Ctrl + Alt + Numpad3
^!Numpad3::SetFillColor("gray")

; Red fill
; Key: Ctrl + Alt + Numpad4
^!Numpad4::SetFillColor("red")

; --- Font Colors ---
; Automatic font color
; Key: Ctrl + Shift + Numpad0
^+Numpad0::SetFontColorAuto()

; Blue font
; Key: Ctrl + Shift + Numpad1
^+Numpad1::SetFontColor("blue")

; Red font
; Key: Ctrl + Shift + Numpad2
^+Numpad2::SetFontColor("red")

; White font
; Key: Ctrl + Shift + Numpad3
^+Numpad3::SetFontColor("white")

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
; Key: Ctrl + Win + Numpad1
^#Numpad1::SetNumberFormat("Accounting")

; General format
; Key: Ctrl + Win + Numpad2
^#Numpad2::SetNumberFormat("General")

; Percentage format
; Key: Ctrl + Win + Numpad3
^#Numpad3::SetNumberFormat("Percentage")

; Date format
; Key: Ctrl + Win + 4
; ^#4::SetNumberFormat("yyyy-mm-dd")

; --- Borders ---
; Remove borders
; Key: Ctrl + Alt + Shift + B
^!+b::SetBorders("Remove")

; Outside borders
; Key: Ctrl + Alt + B
^!b::SetBorders("Outside")

; All borders
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

; Toggle hotstrings globally
#`::ToggleHotstrings()

; ============================================================
; Helper Functions
; ============================================================
; Exits edit mode then sends keystrokes (so they act on the sheet, not a cell)
SendKeys(keys) {
	Send "{Esc}" keys
}

Bind(fn, args*) {
	return (*) => fn(args*)
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
; Fill & Font
; ============================================================
ClearFill() {
	SendKeys("!hhn")
}

SetFillColor(val) {
	SendColor("!hhm", val)
}

SetFontColorAuto() {
	SendKeys("!hfc{Space}")
}

SetFontColor(val) {
	SendColor("!hfcm", val)
}

SendColor(prefix, name) {
	global colors
	if colors.Has(name) {
		rgb := StrReplace(colors[name], ", ", "{Tab}")
		SendKeys(prefix "{Right}{Tab}{Tab}{Tab}{Tab}" rgb "{Enter}")
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

; ============================================================
; Hotstring System
; ============================================================
ToggleHotstrings() {
	global hotstringsEnabled
	hotstringsEnabled := !hotstringsEnabled
	state := hotstringsEnabled ? "ENABLED" : "DISABLED"
	A_IconTip := "Excel Toolkit - Hotstrings " state
	TrayTip "Excel Toolkit", "Hotstrings " state, 1
}

LoadSnippets(file, section) {
	try {
		data := IniRead(file, section)
		if !data
			return Map()
	} catch {
		return Map()
	}

	snippets := Map()

	for rawLine in StrSplit(data, "`n", "`r") {
		line := Trim(rawLine)

		; Skip empty lines and full-line comments
		if !line || SubStr(line, 1, 1) = ";" || SubStr(line, 1, 1) = "#"
			continue

		; Remove inline comments (only if preceded by whitespace)
		line := RegExReplace(line, "(?<=\S)\s+[#].*$")

		if !InStr(line, "=")
			continue

		parts := StrSplit(line, "=", , 2)
		trigger := Trim(parts[1])
		expansion := LTrim(parts[2])

		snippets[trigger] := DecodeEscapes(expansion)
	}

	return snippets
}

DecodeEscapes(str) {
	str := StrReplace(str, "\s", " ")
	str := StrReplace(str, "\S", Chr(0xA0))
	str := StrReplace(str, "\t", "`t")
	str := StrReplace(str, "\n", "`n")
	return str
}

ExpandOrLiteral(trigger, expansion, *) {
	global hotstringsEnabled
	if !hotstringsEnabled {
		Send trigger ";"
		return
	}
	Send expansion
}

; Register hotstrings at startup
if Snippets.Count > 0 {
	HotIf 'WinActive("ahk_exe EXCEL.EXE")'
	for trigger, expansion in Snippets {
		Hotstring(":*:" trigger ";", Bind(ExpandOrLiteral, trigger, expansion))
	}
	HotIf ""
}
