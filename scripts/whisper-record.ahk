#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

; ============================================================
; PATH RESOLUTION — everything is resolved relative to the
; project root so the folder is portable. This script lives at
; <root>/scripts/whisper-record.ahk; rootDir is the parent of
; A_ScriptDir.
; ============================================================
SplitPath(A_ScriptDir, , &parentDir)
global scriptDir := A_ScriptDir
global rootDir := parentDir
global iconsDir := rootDir "\icons"

; ============================================================
; === USER CONFIG — edit these to match your system ==========
; ============================================================
;
; DirectShow audio input name. List your devices with:
;     ffmpeg -list_devices true -f dshow -i dummy
; Use the exact string shown for your mic (the entry marked "(audio)").
global micName := "Microphone (USB audio CODEC)"
;
; Leave empty ("") to auto-detect ffmpeg on PATH + common install locations.
; Set to an absolute path if the auto-detect doesn't find your install.
global ffmpegOverride := ""
;
; Minimum recording length (ms). Shorter releases are silently discarded.
global minRecordMs := 700
;
; Maximum recording length (ms). Hard cap; auto-stops and transcribes.
global maxRecordMs := 90000
;
; Hotkey press/release debounce window (ms).
global debounceMs := 200
;
; Where to anchor the "Copied to clipboard" result toast (the multiline one
; that shows the transcript). The small status toasts — Recording...,
; Transcribing..., errors — always appear bottom-right. Valid values:
; "left"   — bottom-left corner, 20px inset
; "center" — horizontally centered along the bottom edge
; "right"  — bottom-right corner, 20px inset
global resultToastPosition := "left"
;
; ============================================================
; === END USER CONFIG ========================================
; ============================================================

; Derived paths (don't edit — computed from rootDir)
global recordingsDir := rootDir "\recordings"
global transcribeScript := scriptDir "\transcribe.ps1"
global lastTranscriptFile := rootDir "\last-transcript.txt"
global logFile := rootDir "\whisper.log"
global transcribeDebugLog := rootDir "\transcribe-debug.log"
global ffmpegPidFile := rootDir "\ffmpeg.pid"
global iconRec := iconsDir "\rec-1.ico"
global iconCheck := iconsDir "\check-1.ico"
global iconWarn := "shell32.dll"
global iconWarnIdx := 78

; FFREPORT env value (ffmpeg-escaped path); precomputed once.
global ffreportEnvValue := BuildFfreportValue(rootDir "\ffmpeg-report.log")

; ffmpeg executable — auto-detected unless overridden above.
global ffmpeg := ffmpegOverride != "" ? ffmpegOverride : FindFfmpeg()

; Tray icon (stays constant — notification icons are drawn by ShowToast below)
try TraySetIcon(iconsDir "\tray-3.ico")

global currentToast := ""

ShowToast(text, iconFile := "", iconIndex := 0, isError := false, persistent := false, multiline := false, title := "") {
	global currentToast, resultToastPosition

	; Tear down any existing toast so we never stack
	if IsObject(currentToast) {
		try currentToast.Destroy()
		currentToast := ""
	}
	; Also cancel any pending auto-dismiss from a previous toast so it can't
	; fire on top of this new one.
	SetTimer DismissToast, 0

	; -DpiScale keeps every measurement (control sizes, GetPos, Show coords) in
	; raw physical pixels so MonitorGetWorkArea and the window size are in the
	; same unit. Without this, AHK's DPI translation makes GetPos under-report
	; the rendered width and the toast lands off-screen.
	toast := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DpiScale")
	toast.BackColor := isError ? "3a1e1e" : "1e1e1e"

	; Multiline toasts are the "Copied to clipboard" one: box ~1.2x the
	; single-line size, with the transcript body bumped a second ~1.2x
	; on top of that so it's the focal point.
	if multiline {
		toast.MarginX := 38
		toast.MarginY := 31
		toast.SetFont("s13 cFFFFFF", "Segoe UI")
		picOpts := "w48 h48 Section"
	} else {
		toast.MarginX := 32
		toast.MarginY := 26
		toast.SetFont("s11 cFFFFFF", "Segoe UI")
		picOpts := "w40 h40 Section"
	}

	if (iconIndex > 0)
		picOpts .= " Icon" iconIndex

	if (iconFile != "")
		toast.AddPicture(picOpts, iconFile)

	if multiline {
		; Wrapping, auto-size vertically. Icon sits at top; text flows down.
		if (title != "") {
			; Bold header in Segoe UI at the scaled label size, then the
			; transcript body in Georgia (non-italic) at a larger size so the
			; natural leading gives more breathing room between wrapped lines.
			toast.SetFont("s13 cFFFFFF Bold", "Segoe UI")
			toast.AddText("ys x+24 w928", title)
			toast.SetFont("s15 cFFFFFF Norm", "Lora")
			toast.AddText("xp y+16 w928", text)
		} else {
			toast.SetFont("s15 cFFFFFF Norm", "Lora")
			toast.AddText("ys x+24 w928", text)
			toast.SetFont("s13 cFFFFFF Norm", "Segoe UI")
		}
	} else {
		; h40 matches the icon height; +0x200 is SS_CENTERIMAGE which vertically
		; centers the text inside the control, independent of font metrics/DPI.
		toast.AddText("ys x+20 w360 h40 cFFFFFF +0x200", text)
	}

	; Show hidden so AutoSize resolves, measure, then reposition at final spot.
	toast.Show("Hide AutoSize NoActivate")
	toast.GetPos(, , &w, &h)
	MonitorGetWorkArea(, &wLeft, , &wRight, &wBottom)
	; Status toasts always bottom-right; only the multiline result toast
	; honors the resultToastPosition setting.
	if multiline {
		switch resultToastPosition {
			case "left": x := wLeft + 20
			case "center": x := wLeft + ((wRight - wLeft - w) // 2)
			default: x := wRight - w - 20
		}
	} else {
		x := wRight - w - 20
	}
	y := wBottom - h - 20
	toast.Show("x" x " y" y " NoActivate")

	currentToast := toast
	if !persistent {
		; Multiline toasts get longer display time so they can actually be read.
		dismissMs := multiline ? -6000 : -2500
		SetTimer DismissToast, dismissMs
	}
}

DismissToast() {
	global currentToast
	if IsObject(currentToast) {
		try currentToast.Destroy()
		currentToast := ""
	}
}

; Runtime state machine
global state := "idle" ; idle | recording | transcribing
global ffmpegPID := 0
global ffmpegHProcess := 0 ; Windows process HANDLE from CreateProcess; closed after the process exits
global ffmpegHStdin := 0   ; Write end of the stdin pipe; we send "q`n" here for graceful shutdown
global currentOutputFile := ""
global recordStartTick := 0
global lastEventTick := 0
global lastUpTick := 0
global lastTriggerKey := "" ; which hotkey started the current recording

EnsureDir(recordingsDir)
Log("script started; rootDir=" rootDir " ffmpeg=" ffmpeg)
ShowToast("Whisper script loaded", iconRec)

; Clean up an orphaned ffmpeg from a previous crashed run, but ONLY the one
; this script launched. We persist the PID in ffmpegPidFile on start and
; delete it on clean stop; on startup we revive it only if (a) the PID still
; exists and (b) it's actually ffmpeg.exe (defensive PID-reuse guard).
CleanupOrphanFfmpeg()

CleanupOrphanFfmpeg() {
	global ffmpegPidFile
	if !FileExist(ffmpegPidFile)
		return
	try {
		raw := Trim(FileRead(ffmpegPidFile))
		pid := Integer(raw)
		if (pid > 0 && ProcessExist(pid)) {
			name := ""
			try name := ProcessGetName(pid)
			if (name = "ffmpeg.exe") {
				Log("startup: killing leftover ffmpeg pid=" pid " from previous run")
				RunWait('taskkill /F /T /PID ' pid, , "Hide")
			} else {
				Log("startup: stale pid file points to pid=" pid " name=" name " (not ffmpeg); not killing")
			}
		} else {
			Log("startup: pid file references pid=" pid " which no longer exists")
		}
	} catch as err {
		Log("startup: CleanupOrphanFfmpeg exception: " err.Message)
	}
	TryDelete(ffmpegPidFile)
}

; Currently only NumpadAdd (+) is bound. The ergonomic A/B test with
; NumpadEnter is paused — Matt picked + for now. The triggerKey arg is kept
; so adding another binding back is a one-liner.
NumpadAdd:: StartRecording("NumpadAdd")
NumpadAdd Up:: StopRecording("NumpadAdd")

StartRecording(triggerKey := "?", *)
{
	global state, ffmpegPID, ffmpegHProcess, ffmpegHStdin, currentOutputFile, recordStartTick, lastEventTick
	global recordingsDir, micName, ffmpeg, debounceMs, maxRecordMs, ffreportEnvValue
	global lastTriggerKey, ffmpegPidFile, iconRec, iconWarn, iconWarnIdx

	if !DebounceOK()
		return

	if (state != "idle") {
		Log("hotkey down ignored; triggerKey=" triggerKey " state=" state)
		return
	}

	lastTriggerKey := triggerKey

	EnsureDir(recordingsDir)

	stamp := FormatTime(, "yyyyMMdd-HHmmss")
	currentOutputFile := recordingsDir "\recording-" stamp ".wav"
	recordStartTick := A_TickCount

	; -nostats keeps stdout/stderr quiet so the WshShell.Exec pipes don't fill up.
	; -loglevel warning still lets us see errors if anything goes wrong.
	cmd := Format(
		'"{1}" -hide_banner -nostats -loglevel warning -y -f dshow -i audio="{2}" -ar 16000 -ac 1 "{3}"',
		ffmpeg, micName, currentOutputFile
	)

	EnvSet("FFREPORT", ffreportEnvValue)
	Log("launching ffmpeg (Exec): " cmd)

	try {
		; LaunchFfmpegHidden uses CreateProcessW with SW_HIDE so the console
		; window never appears (no flash). It also sets up a stdin pipe we can
		; write "q`n" to later for ffmpeg's graceful shutdown.
		info := LaunchFfmpegHidden(cmd)
		ffmpegHProcess := info["hProcess"]
		ffmpegHStdin := info["hStdin"]
		ffmpegPID := info["pid"]
		state := "recording"
		Log("recording start via " triggerKey "; file=" currentOutputFile " pid=" ffmpegPID)

		; Persist pid so a future script start can clean up an orphan of ours
		try {
			TryDelete(ffmpegPidFile)
			FileAppend(ffmpegPID "", ffmpegPidFile, "UTF-8")
		} catch as err {
			Log("failed to write pid file: " err.Message)
		}

		ShowToast("Recording...", iconRec, 0, false, true)
		SetTimer ForceStopRecording, -maxRecordMs
		SetTimer VerifyFfmpegAlive, -2000
	} catch as err {
		SafeCloseHandle(&ffmpegHProcess)
		SafeCloseHandle(&ffmpegHStdin)
		ffmpegPID := 0
		currentOutputFile := ""
		state := "idle"
		lastTriggerKey := ""
		Log("recording failed to start via " triggerKey ": " err.Message)
		ShowToast("Could not start recording", iconWarn, iconWarnIdx, true)
	}
}

VerifyFfmpegAlive() {
	global state, ffmpegPID, ffmpegHProcess, ffmpegHStdin, currentOutputFile, lastTriggerKey
	global ffmpegPidFile, iconWarn, iconWarnIdx

	if (state != "recording")
		return

	if !ProcessExist(ffmpegPID) {
		Log("WARNING: ffmpeg pid=" ffmpegPID " DIED within 2000ms of start; output file=" currentOutputFile)
		; Roll the script back to a clean idle state so the next hotkey press works.
		SetTimer ForceStopRecording, 0
		; Brief settle so any lingering OS handle on the zombie wav is released
		; before we try to delete it (reduces transient TryDelete failures).
		Sleep 150
		TryDelete(currentOutputFile)
		TryDelete(ffmpegPidFile)
		SafeCloseHandle(&ffmpegHProcess)
		SafeCloseHandle(&ffmpegHStdin)
		ffmpegPID := 0
		currentOutputFile := ""
		state := "idle"
		lastTriggerKey := ""
		ShowToast("Recording failed to start", iconWarn, iconWarnIdx, true)
		return
	}

	try {
		size := FileExist(currentOutputFile) ? FileGetSize(currentOutputFile) : -1
		Log("ffmpeg alive check OK; pid=" ffmpegPID " fileSize=" size)
	} catch as err {
		Log("ffmpeg alive check: FileGetSize failed: " err.Message)
	}
}

; LaunchFfmpegHidden: launches a command with a hidden window AND a redirected
; stdin pipe so we can later send "q" for a graceful ffmpeg shutdown.
;
; Why this exists: AHK's Run(..., "Hide") passes CREATE_NO_WINDOW which gives
; the child process no console at all — that broke ffmpeg's dshow audio capture
; earlier. WshShell.Exec gives us a stdin pipe but the console window is always
; visible (brief flash even if we WinHide it later). CreateProcess with
; STARTF_USESHOWWINDOW + SW_HIDE (and NO CREATE_NO_WINDOW) hides the window
; from the first frame while keeping a real console attached — the combo that
; keeps dshow happy with zero visible popup.
;
; Returns a Map with:
;   hProcess - Windows process HANDLE (close with CloseHandle when done)
;   pid      - ffmpeg's process ID
;   hStdin   - write end of stdin pipe (close with CloseHandle when done)
LaunchFfmpegHidden(cmd) {
	; SECURITY_ATTRIBUTES { nLength=24, lpSecurityDescriptor=NULL, bInheritHandle=TRUE }
	sa := Buffer(24, 0)
	NumPut("UInt", 24, sa, 0)
	NumPut("Int", 1, sa, 16)

	; CreatePipe(&hRead, &hWrite, &sa, 0)
	hReadPipe := 0, hWritePipe := 0
	if !DllCall("kernel32\CreatePipe", "Ptr*", &hReadPipe, "Ptr*", &hWritePipe, "Ptr", sa.Ptr, "UInt", 0)
		throw Error("CreatePipe failed, err=" A_LastError)

	; Make the WRITE end non-inheritable so only the read end goes into the child.
	; HANDLE_FLAG_INHERIT = 0x1
	DllCall("kernel32\SetHandleInformation", "Ptr", hWritePipe, "UInt", 1, "UInt", 0)

	; Open NUL for stdout/stderr so the child's output is silently discarded
	; (must be inheritable so the child can use it).
	hNul := DllCall("kernel32\CreateFileW",
		"Str", "NUL",
		"UInt", 0x40000000,       ; GENERIC_WRITE
		"UInt", 0x3,              ; FILE_SHARE_READ | FILE_SHARE_WRITE
		"Ptr", sa.Ptr,            ; inheritable
		"UInt", 3,                ; OPEN_EXISTING
		"UInt", 0, "Ptr", 0, "Ptr")
	if (hNul = -1 || hNul = 0) {
		DllCall("CloseHandle", "Ptr", hReadPipe)
		DllCall("CloseHandle", "Ptr", hWritePipe)
		throw Error("CreateFile NUL failed, err=" A_LastError)
	}

	; STARTUPINFOW (104 bytes on x64)
	si := Buffer(104, 0)
	NumPut("UInt", 104, si, 0)                  ; cb
	NumPut("UInt", 0x101, si, 60)               ; dwFlags = STARTF_USESTDHANDLES (0x100) | STARTF_USESHOWWINDOW (0x1)
	NumPut("UShort", 0, si, 64)                 ; wShowWindow = SW_HIDE
	NumPut("Ptr", hReadPipe, si, 80)            ; hStdInput
	NumPut("Ptr", hNul, si, 88)                 ; hStdOutput
	NumPut("Ptr", hNul, si, 96)                 ; hStdError

	; PROCESS_INFORMATION (24 bytes on x64)
	pi := Buffer(24, 0)

	; lpCommandLine must be a writable UTF-16 buffer
	cmdBuf := Buffer((StrLen(cmd) + 1) * 2, 0)
	StrPut(cmd, cmdBuf, "UTF-16")

	; CreateProcessW with bInheritHandles=TRUE and dwCreationFlags=0
	; (no CREATE_NO_WINDOW — that's what broke dshow before)
	ok := DllCall("kernel32\CreateProcessW",
		"Ptr", 0,
		"Ptr", cmdBuf.Ptr,
		"Ptr", 0,
		"Ptr", 0,
		"Int", 1,                 ; bInheritHandles
		"UInt", 0,                ; dwCreationFlags
		"Ptr", 0,
		"Ptr", 0,
		"Ptr", si.Ptr,
		"Ptr", pi.Ptr)

	; Close our copies of the inheritable handles; the child has its own.
	DllCall("CloseHandle", "Ptr", hReadPipe)
	DllCall("CloseHandle", "Ptr", hNul)

	if !ok {
		DllCall("CloseHandle", "Ptr", hWritePipe)
		throw Error("CreateProcess failed, err=" A_LastError)
	}

	hProcess := NumGet(pi, 0, "Ptr")
	hThread := NumGet(pi, 8, "Ptr")
	pid := NumGet(pi, 16, "UInt")

	; We don't need the thread handle
	DllCall("CloseHandle", "Ptr", hThread)

	return Map("hProcess", hProcess, "pid", pid, "hStdin", hWritePipe)
}

; Write "q`n" to a pipe handle (used to tell ffmpeg to stop cleanly)
WritePipeString(handle, text) {
	if !handle
		return 0
	bufSize := StrPut(text, "UTF-8")   ; includes null terminator
	buf := Buffer(bufSize, 0)
	bytesToWrite := StrPut(text, buf, "UTF-8") - 1  ; exclude null
	written := 0
	DllCall("kernel32\WriteFile",
		"Ptr", handle,
		"Ptr", buf.Ptr,
		"UInt", bytesToWrite,
		"UInt*", &written,
		"Ptr", 0)
	return written
}

; Query whether a process HANDLE is still running. Returns true if still alive.
IsProcessAlive(hProcess) {
	if !hProcess
		return false
	; WaitForSingleObject with 0 timeout: WAIT_TIMEOUT (0x102) means still running
	result := DllCall("kernel32\WaitForSingleObject", "Ptr", hProcess, "UInt", 0)
	return (result = 0x102)
}

; Build the value for the FFREPORT env var. ffmpeg's FFREPORT parser treats
; unescaped ':' as an option separator, so the drive-letter colon has to be
; escaped as '\:'. Using forward slashes keeps the rest of the path clean.
BuildFfreportValue(logPath) {
	s := StrReplace(logPath, "\", "/")
	s := StrReplace(s, ":", "\:")
	return "file=" s ":level=48"
}

; Locate ffmpeg.exe. Searches PATH first, then falls back to a small list of
; common install locations. If nothing is found we show a friendly error and
; exit rather than throwing a stack trace at the user.
FindFfmpeg() {
	pathEnv := EnvGet("PATH")
	for dir in StrSplit(pathEnv, ";") {
		if (dir = "")
			continue
		candidate := RTrim(dir, "\") "\ffmpeg.exe"
		if FileExist(candidate)
			return candidate
	}
	fallbacks := [
		"C:\ProgramData\chocolatey\bin\ffmpeg.exe",
		"C:\Program Files\ffmpeg\bin\ffmpeg.exe",
		"C:\ffmpeg\bin\ffmpeg.exe",
	]
	for candidate in fallbacks {
		if FileExist(candidate)
			return candidate
	}
	MsgBox(
		"ffmpeg.exe was not found on PATH or in any common install location.`n`n"
		"Install ffmpeg (for example: choco install ffmpeg) or set ``ffmpegOverride``"
		" at the top of the script to the full path to your ffmpeg.exe.",
		"Whisper script", "IconX"
	)
	ExitApp
}

SafeCloseHandle(&handle) {
	if handle {
		try DllCall("kernel32\CloseHandle", "Ptr", handle)
		handle := 0
	}
}

StopRecording(triggerKey := "?", *)
{
	global state, ffmpegPID, currentOutputFile, recordStartTick
	global minRecordMs, lastUpTick, debounceMs, lastTriggerKey, iconRec

	now := A_TickCount
	if ((now - lastUpTick) < debounceMs) {
		Log("hotkey up debounced; triggerKey=" triggerKey " deltaMs=" (now - lastUpTick))
		return
	}
	lastUpTick := now

	if (state != "recording") {
		Log("hotkey up ignored; triggerKey=" triggerKey " state=" state)
		return
	}

	recordDuration := A_TickCount - recordStartTick
	keyStatus := (triggerKey = lastTriggerKey) ? "match" : "MISMATCH(start=" lastTriggerKey ")"
	Log("recording stop via " triggerKey " " keyStatus "; durationMs=" recordDuration " pid=" ffmpegPID " file=" currentOutputFile)

	SetTimer ForceStopRecording, 0
	Log("force-stop timer cancelled")

	StopFfmpeg(ffmpegPID)
	ffmpegPID := 0

	LogFileState(currentOutputFile, "post-stop")

	if (recordDuration < minRecordMs) {
		Log("recording too short (" recordDuration "ms < " minRecordMs "ms); deleting")
		Sleep 300
		TryDelete(currentOutputFile)
		currentOutputFile := ""
		state := "idle"
		lastTriggerKey := ""
		ShowToast("Recording too short, skipped", iconRec)
		return
	}

	fileToTranscribe := currentOutputFile
	currentOutputFile := ""
	TranscribeAndPaste(fileToTranscribe)
}

ForceStopRecording() {
	global state, ffmpegPID, currentOutputFile

	if (state != "recording") {
		Log("force-stop fired but state=" state "; ignoring")
		return
	}

	SetTimer ForceStopRecording, 0
	Log("force-stop triggered (max time hit); pid=" ffmpegPID " file=" currentOutputFile)

	StopFfmpeg(ffmpegPID)
	ffmpegPID := 0

	LogFileState(currentOutputFile, "post-stop (force)")

	fileToTranscribe := currentOutputFile
	currentOutputFile := ""
	TranscribeAndPaste(fileToTranscribe)
}

TranscribeAndPaste(filePath) {
	global state, transcribeScript, lastTranscriptFile, lastTriggerKey, transcribeDebugLog
	global iconRec, iconCheck, iconWarn, iconWarnIdx

	if (state = "transcribing") {
		Log("TranscribeAndPaste re-entry blocked; state already=transcribing")
		return
	}

	state := "transcribing"
	ShowToast("Transcribing...", iconRec)
	Log("=== TranscribeAndPaste start; file=" filePath " ===")

	try {
		LogFileState(filePath, "pre-settle")
		Sleep 300
		LogFileState(filePath, "post-settle")

		if !FileExist(filePath) {
			Log("recording file missing after settle: " filePath)
			ShowToast("Recording file not found", iconWarn, iconWarnIdx, true)
			return
		}

		fileSize := FileGetSize(filePath)
		if (fileSize = 0) {
			Log("WARNING: wav file is 0 bytes; skipping transcription")
			ShowToast("Recording file is empty", iconWarn, iconWarnIdx, true)
			return
		}

		psCmd := 'cmd.exe /c powershell.exe -ExecutionPolicy Bypass -File "' transcribeScript '" -InputFile "' filePath '" >> "' transcribeDebugLog '" 2>&1'
		Log("invoking transcribe: " psCmd)

		try {
			psStart := A_TickCount
			exitCode := RunWait(psCmd, , "Hide")
			psDuration := A_TickCount - psStart
			Log("transcribe finished; exitCode=" exitCode " durationMs=" psDuration)
		} catch as err {
			Log("transcribe RunWait threw: " err.Message)
			ShowToast("Transcription failed", iconWarn, iconWarnIdx, true)
			return
		}

		txtFile := filePath ".txt"
		if !FileExist(txtFile) {
			fallback := RegExReplace(filePath, "\.wav$", ".txt")
			Log("primary transcript path missing (" txtFile "); trying fallback " fallback)
			if FileExist(fallback)
				txtFile := fallback
		}
		Log("looking for transcript at: " txtFile)

		if FileExist(txtFile) {
			try {
				if FileExist(lastTranscriptFile)
					FileDelete(lastTranscriptFile)
			}
			try FileCopy(txtFile, lastTranscriptFile, true)
			try {
				transcript := Trim(FileRead(txtFile, "UTF-8"))
				A_Clipboard := transcript
				Log("clipboard updated from " txtFile " (" StrLen(transcript) " chars)")
				; Show the transcript in the toast so Matt can glance at what got captured.
				; Truncate very long ones so the toast doesn't take over the screen —
				; the full text is still on the clipboard regardless.
				maxToastLen := 400
				displayText := (StrLen(transcript) > maxToastLen)
					? SubStr(transcript, 1, maxToastLen) "…"
					: transcript
				ShowToast(displayText, iconCheck, 0, false, false, true, "Copied to clipboard")
			} catch as err {
				Log("clipboard update failed: " err.Message)
				ShowToast("Transcript saved, clipboard failed", iconWarn, iconWarnIdx, true)
			}
		} else {
			Log("transcript file missing: " txtFile)
			ShowToast("Transcript not found", iconWarn, iconWarnIdx, true)
		}
	} finally {
		state := "idle"
		lastTriggerKey := ""
		Log("=== TranscribeAndPaste end; state=idle ===")
	}
}

StopFfmpeg(pid) {
	global ffmpegHProcess, ffmpegHStdin, currentOutputFile, ffmpegPidFile

	if !pid {
		Log("StopFfmpeg: no pid given")
		return
	}

	Log("StopFfmpeg: entry pid=" pid " processExist=" (ProcessExist(pid) ? "yes" : "no") " hasStdin=" (ffmpegHStdin ? "yes" : "no"))
	LogFileState(currentOutputFile, "StopFfmpeg entry")

	; --- Graceful: write "q`n" to ffmpeg stdin pipe ---
	gracefulOK := false
	if ffmpegHStdin {
		try {
			written := WritePipeString(ffmpegHStdin, "q`n")
			Log("StopFfmpeg: 'q' written to stdin; bytes=" written)

			; Close our end of stdin so ffmpeg sees EOF — helps it exit promptly.
			SafeCloseHandle(&ffmpegHStdin)

			waitStart := A_TickCount
			lastLoggedSize := -1
			while (IsProcessAlive(ffmpegHProcess) && (A_TickCount - waitStart) < 5000) {
				Sleep 100
				try {
					curSize := FileExist(currentOutputFile) ? FileGetSize(currentOutputFile) : -1
				} catch {
					curSize := -2
				}
				if (curSize != lastLoggedSize) {
					Log("StopFfmpeg: waiting... elapsed=" (A_TickCount - waitStart) "ms fileSize=" curSize)
					lastLoggedSize := curSize
				}
			}
			waitMs := A_TickCount - waitStart

			if !IsProcessAlive(ffmpegHProcess) {
				exitCode := 0
				try DllCall("kernel32\GetExitCodeProcess", "Ptr", ffmpegHProcess, "UInt*", &exitCode)
				Log("StopFfmpeg: graceful exit OK waitMs=" waitMs " exitCode=" exitCode)
				gracefulOK := true
			} else {
				Log("StopFfmpeg: graceful wait TIMEOUT after " waitMs "ms; ffmpeg still running")
			}
		} catch as err {
			Log("StopFfmpeg: stdin-write exception: " err.Message)
		}
	} else {
		Log("StopFfmpeg: no stdin pipe available; going straight to taskkill")
	}

	; --- Fallback: taskkill ---
	if !gracefulOK {
		if ProcessExist(pid) {
			Log("StopFfmpeg: fallback taskkill /F /T /PID " pid)
			try {
				tkExit := RunWait('taskkill /F /T /PID ' pid, , "Hide")
				Log("StopFfmpeg: taskkill exitCode=" tkExit)
			} catch as err {
				Log("StopFfmpeg: taskkill exception: " err.Message)
			}
		} else {
			Log("StopFfmpeg: pid=" pid " already gone before taskkill")
		}
	}

	; Let the OS flush the output file / release the device
	Sleep 200
	SafeCloseHandle(&ffmpegHStdin)
	SafeCloseHandle(&ffmpegHProcess)
	TryDelete(ffmpegPidFile)

	LogFileState(currentOutputFile, "StopFfmpeg exit")
	if ProcessExist(pid)
		Log("StopFfmpeg: WARNING pid=" pid " STILL alive at exit")
	else
		Log("StopFfmpeg: pid=" pid " confirmed gone")
}

LogFileState(path, label) {
	if !path {
		Log(label ": (no path)")
		return
	}
	if !FileExist(path) {
		Log(label ": file does not exist: " path)
		return
	}
	try {
		size := FileGetSize(path)
		Log(label ": size=" size " bytes path=" path)
	} catch as err {
		Log(label ": FileGetSize failed: " err.Message)
	}
}

DebounceOK() {
	global lastEventTick, debounceMs

	now := A_TickCount
	if ((now - lastEventTick) < debounceMs)
		return false

	lastEventTick := now
	return true
}

EnsureDir(path) {
	if !DirExist(path)
		DirCreate(path)
}

TryDelete(path) {
	if !path
		return
	if !FileExist(path)
		return
	try {
		FileDelete(path)
	} catch as err {
		Log("TryDelete failed for " path ": " err.Message)
	}
}

Log(msg) {
	global logFile
	line := FormatTime(, "yyyy-MM-dd HH:mm:ss") " | " msg "`n"
	FileAppend(line, logFile, "UTF-8")
}
