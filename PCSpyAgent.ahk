; ============================================================
; PC Activity Agent - Sends Events to MacroDroid via Webhook
; Created by: A Dev Named DeLL
; Purpose: Track user activity & send updates to phone
; ============================================================

#SingleInstance, Force
#NoTrayIcon
#Persistent

; Your MacroDroid webhook URL
webhookLink := "https://trigger.macrodroid.com/a44fce57-a3ac-4418-861c-363d67987378/spymypc"

; Core runtime variables
lastWindowTitle := ""
lastClipboard := ""
lastSentTimestamp := 0
logFolder := A_AppData "\PCSpyAgent"
logFilePath := logFolder "\loggs.txt"

; File change tracking variables
trackedFiles := {}
currentFiles := {}

; Initialize timers and message hooks
SetTimer, MonitorActiveWindow, 1000
SetTimer, MonitorClipboardChange, 1500
SetTimer, MonitorIdleTime, 5000
SetTimer, MonitorFileChanges, 3000
OnMessage(0x219, "DetectUSBChange")
OnMessage(0x218, "DetectSessionEvent")
return

; =========================
; Active Window Monitor
; =========================
MonitorActiveWindow:
    WinGetTitle, activeTitle, A
    if (activeTitle != lastWindowTitle) {
        lastWindowTitle := activeTitle

        WinGet, activeProcess, ProcessName, A
        WinGet, activeHwnd, ID, A
        WinGetClass, activeClass, A

        if (activeProcess = "explorer.exe") {
            for shellWin in ComObjCreate("Shell.Application").Windows {
                try {
                    if (InStr(shellWin.FullName, "explorer.exe") && shellWin.HWND = activeHwnd) {
                        folderPath := ""
                        try folderPath := shellWin.Document.Folder.Self.Path
                        readablePath := FormatVirtualFolder(folderPath)
                        SendToPhone("File Explorer", readablePath)
                        return
                    }
                }
            }
        }
        else if (activeProcess = "chrome.exe" && activeClass = "Chrome_WidgetWin_1") {
            SendToPhone("Google Chrome", activeTitle)
        }
        else {
            SendToPhone("Window Change", activeTitle)
        }
    }
return

; =====================
; Clipboard Monitor
; =====================
MonitorClipboardChange:
    if (Clipboard != lastClipboard) {
        lastClipboard := Clipboard
        shortClip := SubStr(lastClipboard, 1, 200)
        SendToPhone("Clipboard Updated", shortClip)
    }
return

; =====================
; USB Plug/Unplug Notification
; =====================
DetectUSBChange(wParam, lParam) {
    if (wParam = 0x8000)
        SendToPhone("USB Event", "Device plugged in")
    else if (wParam = 0x8004)
        SendToPhone("USB Event", "Device removed")
}

; =====================
; Session Lock/Unlock Detection
; =====================
DetectSessionEvent(wParam, lParam) {
    if (wParam = 0x7)
        SendToPhone("Session Status", "User locked the session")
    else if (wParam = 0x8)
        SendToPhone("Session Status", "User unlocked the session")
}

; =====================
; Idle Time Monitor
; =====================
MonitorIdleTime:
    idleMinutes := A_TimeIdle / 60000
    if (idleMinutes >= 5) {
        SendToPhone("Idle Detected", "User idle for " . Round(idleMinutes, 1) . " min")
    }
return

; ==========================
; File Creation / Deletion Monitor
; ==========================
MonitorFileChanges:
    folderToMonitor := A_Desktop ; Change if you want to monitor other folders

    ; Get current files and timestamps in monitored folder
    newFiles := {}
    Loop, Files, % folderToMonitor "\*.*", FR  ; Include files, recurse false
    {
        newFiles[A_LoopFileFullPath] := A_LoopFileTime
    }

    ; Detect created files
    for filePath, fileTime in newFiles {
        if (!currentFiles.HasKey(filePath)) {
            SendToPhone("File Created", "File: " . SubStr(filePath, StrLen(folderToMonitor)+2))
        }
        else if (currentFiles[filePath] != fileTime) {
            SendToPhone("File Modified", "File: " . SubStr(filePath, StrLen(folderToMonitor)+2))
        }
    }

    ; Detect deleted files
    for filePath, _ in currentFiles {
        if (!newFiles.HasKey(filePath)) {
            SendToPhone("File Deleted", "File: " . SubStr(filePath, StrLen(folderToMonitor)+2))
        }
    }

    ; Update currentFiles with new snapshot
    currentFiles := newFiles
return

; ==========================
; Send Data to MacroDroid
; ==========================
SendToPhone(eventType, eventDetail) {
    global webhookLink, lastSentTimestamp, logFilePath

    if (A_TickCount - lastSentTimestamp < 1000)  ; rate limit (1s)
        return
    lastSentTimestamp := A_TickCount

    encodedType := URLEncode(eventType)
    encodedDetail := URLEncode(eventDetail)
    timestamp := URLEncode(A_Now)

    finalURL := webhookLink . "?type=" . encodedType . "&detail=" . encodedDetail . "&timestamp=" . timestamp

    try {
        http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", finalURL)
        http.Send()
        LogNotification(eventType, eventDetail)
    } catch {
        ; Fail silently
    }
}

; ==========================
; Log Notifications to File
; ==========================
LogNotification(eventType, eventDetail) {
    global logFolder, logFilePath
    if !FileExist(logFolder) {
        FileCreateDir, % logFolder
    }
    FileAppend, % A_Now . " - " . eventType . ": " . eventDetail . "`n", %logFilePath%
}

; ==========================
; URL-Safe Encoding Function
; ==========================
URLEncode(text) {
    static safeChars := "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.~"
    encoded := ""
    Loop, Parse, text
    {
        ch := A_LoopField
        if InStr(safeChars, ch)
            encoded .= ch
        else
            encoded .= "%" . Format("{:02X}", Asc(ch))
    }
    return encoded
}

; ==========================
; Normalize Virtual Folder Path
; ==========================
FormatVirtualFolder(folderPath) {
    StringLower, cleanPath, folderPath
    cleanPath := Trim(cleanPath)
    RegExMatch(cleanPath, "\{[0-9a-f\-]{36}\}", foundCLSID)

    if (foundCLSID) {
        RegRead, virtualName, HKEY_CLASSES_ROOT\CLSID\%foundCLSID%,
        if (virtualName != "")
            return virtualName
    }
    return folderPath
}
