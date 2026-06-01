#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

InstallKeybdHook()
InstallMouseHook()

; Live capture: every key DOWN fires a tray tip + tooltip showing the
; exact name + vk/sc so you can read it without opening Key History.
; Press the shutter button and watch for a tooltip near the cursor.
~*$Joy1::ShowKey("Joy1")  ; some BT shutters present as a gamepad

F12::KeyHistory
Esc::ExitApp

; Catch-all keyboard hook via InputHook — logs whatever comes through.
ih := InputHook("L0 V")
ih.KeyOpt("{All}", "N")
ih.OnKeyDown := OnAnyKey
ih.Start()

OnAnyKey(hook, vk, sc) {
	name := GetKeyName(Format("vk{:X}sc{:X}", vk, sc))
	combo := Format("vk{:X}sc{:X}", vk, sc)
	ShowKey(name " (" combo ")")
}

ShowKey(txt) {
	ToolTip("KEY: " txt)
	SetTimer(() => ToolTip(), -2500)
}
