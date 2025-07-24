#Requires AutoHotkey v2.0
#SingleInstance Force

;========================================================================================================
;
;                                 WinMacros - fr0st
;
;   Want to customize your script? Feel free to make adjustments!
;
;  # ISSUES
;
;   Have an idea, suggestion, or an issue? You can share it by creating an issue here:
;   https://github.com/fr0st-iwnl/WinMacros/issues
;
;  # PULL REQUESTS
;
;   If you'd like to contribute or add something to the script, submit a pull request here:
;   https://github.com/fr0st-iwnl/WinMacros/pulls
;
;========================================================================================================

; Class for applying dark mode to system tray and popup menus
; big thanks to NPerovic / https://www.autohotkey.com/boards/viewtopic.php?t=114808
Class darkMode
{
    ; Mode: Dark = 1, Default (Light) = 0   
    Static SetMode(Mode := 1) {
        DllCall(DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", 135, "ptr"), "int", Mode)
        DllCall(DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", 136, "ptr"))
    }
}

; Fade in animation
FadeIn(hwnd) {
    Loop 10 {
        opacity := A_Index * 25
        WinSetTransparent(Integer(opacity), "ahk_id " . hwnd)
        Sleep(15)
    }
}

; Fade out animation
FadeOut(hwnd) {
    Loop 10 {
        opacity := 250 - (A_Index * 25)
        WinSetTransparent(Integer(opacity), "ahk_id " . hwnd)
        Sleep(20)
    }

    WinSetTransparent(0, "ahk_id " . hwnd)
    Sleep(50)
}

; my god i fucking hate making GUIs in AHK

SafeStyleButton(button, isDarkTheme := true) {
    ; Apply styling without using the problematic SetColor method
    if (isDarkTheme) {
        ; Dark theme styling
        button.SetFont("c0xFFFFFF")
        button.Opt("+Background333333")
    } else {
        ; Light theme styling - make buttons lighter
        button.SetFont("c0x000000") 
        button.Opt("+BackgroundF0F0F0")
    }
    
    ; Add rounded corners using the Windows API directly
    if (VerCompare(A_OSVersion, "10.0.22200") >= 0) { ; Windows 11
        try {
            ; Apply appropriate theme based on dark/light mode
            if (isDarkTheme) {
                DllCall("uxtheme\SetWindowTheme", "Ptr", button.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
            } else {
                DllCall("uxtheme\SetWindowTheme", "Ptr", button.Hwnd, "Str", "Explorer", "Ptr", 0)
            }
            
            hwnd := button.Hwnd
            rc := Buffer(16)
            DllCall("User32\GetWindowRect", "Ptr", hwnd, "Ptr", rc)
            width := NumGet(rc, 8, "Int") - NumGet(rc, 0, "Int")
            height := NumGet(rc, 12, "Int") - NumGet(rc, 4, "Int")
            region := DllCall("Gdi32\CreateRoundRectRgn", "Int", 0, "Int", 0, "Int", width, "Int", height, "Int", 9, "Int", 9, "Ptr")
            DllCall("User32\SetWindowRgn", "Ptr", hwnd, "Ptr", region, "Int", 1)
            DllCall("Gdi32\DeleteObject", "Ptr", region)
        }
    }
    
    return button
}

; some variables
global keybindsFile := EnvGet("LOCALAPPDATA") "\WinMacros\keybinds.ini"
global isMenuOpen := false
global hotkeyActions := Map(
    "OpenExplorer", "Open File Explorer",
    "OpenPowerShell", "Open PowerShell",
    "ToggleMic", "Toggle Microphone",
    "OpenBrowser", "Open Default Browser",
    "VolumeUp", "Volume Up",
    "VolumeDown", "Volume Down",
    "ToggleMute", "Toggle Volume Mute",
    "ToggleTaskbar", "Toggle Taskbar",
    "ToggleDesktopIcons", "Toggle Desktop Icons",
    "OpenVSCode", "Open Code Editor",
    "OpenCalculator", "Open Calculator",
    "OpenSpotify", "Open Spotify"
)

global settingsFile := EnvGet("LOCALAPPDATA") "\WinMacros\settings.ini"

global currentVersion := "1.6"
global versionCheckUrl := "https://winmacros.netlify.app/version/version.txt"
global githubReleasesUrl := "https://github.com/fr0st-iwnl/WinMacros/releases/latest"

global currentTheme := IniRead(settingsFile, "Settings", "Theme", "dark")
global notificationsEnabled := IniRead(settingsFile, "Settings", "Notifications", "1") = "1"
global animationsEnabled := IniRead(settingsFile, "Settings", "Animations", "0") = "1"

global unifiedGui := ""
global currentTabIndex := 1

global currentSetHotkeyGui := ""
global currentSetHotkeyAction := ""
global isHotkeyGuiOpen := false

global powerGui := ""
global isLauncherEditGuiOpen := false

global isCheckingForUpdates := false

global isEditorMenuOpen := false
global editorGui := ""
global currentSelection := 1

global myGui := ""

global startupChk := ""

global activeNotifications := []

global notificationQueue := []
global isShowingNotification := false
global currentNotify := ""

global launcherIniPath := EnvGet("LOCALAPPDATA") "\WinMacros\launcher.ini"
global activeHotkeys := Map()

if !DirExist(EnvGet("LOCALAPPDATA") "\WinMacros") {
    DirCreate(EnvGet("LOCALAPPDATA") "\WinMacros")
}

if !FileExist(keybindsFile) {
    for action, description in hotkeyActions {
        IniWrite("None", keybindsFile, "Hotkeys", action)
    }
}

if !FileExist(settingsFile) {
    IniWrite(1, settingsFile, "Settings", "ShowWelcome")
    IniWrite("dark", settingsFile, "Settings", "Theme")
    IniWrite(1, settingsFile, "Settings", "Notifications")
    IniWrite(0, settingsFile, "Settings", "Animations")
}

global currentTheme := IniRead(settingsFile, "Settings", "Theme", "dark")
global notificationsEnabled := IniRead(settingsFile, "Settings", "Notifications", "1") = "1"

if (currentTheme = "dark") {
    darkMode.SetMode(1)
}

; Function to hide focus borders in GUI controls [ex: unfocused textboxes, etc]
; still kinda buggy, but it works for now [press TAB in the GUI to see it in action]
HideFocusBorder(hWnd) {
    ; WM_UPDATEUISTATE = 0x0128
    ; UIS_SET << 16 | UISF_HIDEFOCUS = 0x00010001
    static HideFocus := 0x00010001
    
    if DllCall("IsWindow", "Ptr", hWnd, "UInt")
        PostMessage(0x0128, HideFocus, 0, , "ahk_id " hWnd)
}

OnMessage(0x0128, WM_UPDATEUISTATE)
WM_UPDATEUISTATE(wParam, lParam, msg, hWnd) {
    static HideFocus := 0x00010001
    static Affected := Map()
    
    if (wParam = HideFocus)
        Affected[hWnd] := true
    else if Affected.Has(hWnd)
        PostMessage(0x0128, HideFocus, 0, , "ahk_id " hWnd)
}

InitializeTrayMenu() {
    A_IconTip := "WinMacros"

    A_TrayMenu.Delete()
    
    A_TrayMenu.Add("Open WinMacros", (*) => ShowUnifiedGUI(1))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Show Welcome Message at startup", ToggleStartup)
    A_TrayMenu.Add("Run on Windows Startup", ToggleWindowsStartup)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Dark Theme", ToggleTrayTheme)
    A_TrayMenu.Add("Enable Notifications", ToggleNotifications)
    A_TrayMenu.Add("Enable Animations", ToggleAnimations)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Check for Updates", (*) => CheckForUpdates(true))
    A_TrayMenu.Add("Exit", SafeExitApp)

    ; Set default item to open WinMacros GUI and enable single click
    A_TrayMenu.Default := "Open WinMacros"
    A_TrayMenu.ClickCount := 1
    
    if (currentTheme = "dark") {
        A_TrayMenu.Check("Dark Theme")
        darkMode.SetMode(1)
    } else {
        A_TrayMenu.Uncheck("Dark Theme")
        darkMode.SetMode(0)
    }
    
    if (IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1")
        A_TrayMenu.Check("Show Welcome Message at startup")
    if (FileExist(A_Startup "\WinMacros.lnk"))
        A_TrayMenu.Check("Run on Windows Startup")
    if (notificationsEnabled)
        A_TrayMenu.Check("Enable Notifications")
    if (animationsEnabled)
        A_TrayMenu.Check("Enable Animations")
}

; Safe exit function to prevent errors when update message box is open
SafeExitApp(*) {
    global isCheckingForUpdates
    
    ; Check if an update message box is active
    if (isCheckingForUpdates && WinExist("Update Available") || WinExist("Update Check Failed") || WinExist("No Updates Available")) {
        ShowNotification("❌ Please close the update dialog before exiting")
        return
    }
    
    ExitApp()
}

; here we go bam bam bam
InitializeTrayMenu()

ToggleStartup(*) {
    global settingsFile, unifiedGui
    isChecked := IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1"
    
    newState := !isChecked
    IniWrite(newState ? "1" : "0", settingsFile, "Settings", "ShowWelcome")
    
    if (newState) {
        ShowNotification("✅ Welcome screen enabled at startup")
        A_TrayMenu.Check("Show Welcome Message at startup")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["ShowWelcome"]) {
                    unifiedGui["ShowWelcome"].Value := true
                }
            }
        }
    } else {
        ShowNotification("❌ Welcome screen disabled at startup")
        A_TrayMenu.Uncheck("Show Welcome Message at startup")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["ShowWelcome"]) {
                    unifiedGui["ShowWelcome"].Value := false
                }
            }
        }
    }
}

ToggleWindowsStartup(*) {
    global unifiedGui
    startupPath := A_Startup "\WinMacros.lnk"
    
    if (!FileExist(startupPath)) {
        try {
            FileCreateShortcut(A_ScriptFullPath, startupPath,, "Launch WinMacros on startup")
            ShowNotification("✅ WinMacros will run on Windows startup")
            A_TrayMenu.Check("Run on Windows Startup")
            
            if (IsObject(unifiedGui) && WinExist("WinMacros")) {
                try {
                    if (unifiedGui["RunOnStartup"]) {
                        unifiedGui["RunOnStartup"].Value := true
                    }
                }
            }
        } catch Error as err {
            ShowNotification("❌ Failed to create startup shortcut")
        }
    } else {
        try {
            FileDelete(startupPath)
            ShowNotification("❌ WinMacros will not run on Windows startup")
            A_TrayMenu.Uncheck("Run on Windows Startup")
            
            if (IsObject(unifiedGui) && WinExist("WinMacros")) {
                try {
                    if (unifiedGui["RunOnStartup"]) {
                        unifiedGui["RunOnStartup"].Value := false
                    }
                }
            }
        } catch Error as err {
            ShowNotification("❌ Failed to remove startup shortcut")
        }
    }
}

ToggleNotifications(*) {
    global notificationsEnabled, settingsFile, unifiedGui
    
    notificationsEnabled := !notificationsEnabled
    IniWrite(notificationsEnabled ? "1" : "0", settingsFile, "Settings", "Notifications")
    
    if (notificationsEnabled) {
        A_TrayMenu.Check("Enable Notifications")
        ShowNotification("📢 Notifications enabled")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["EnableNotifications"]) {
                    unifiedGui["EnableNotifications"].Value := true
                }
            }
        }
    } else {
        A_TrayMenu.Uncheck("Enable Notifications")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["EnableNotifications"]) {
                    unifiedGui["EnableNotifications"].Value := false
                }
            }
        }
    }
}

ToggleTrayTheme(*) {
    global currentTheme
    ToggleTheme(currentTheme = "light")
}

CreateTrayMenu() {
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Open WinMacros", (*) => ShowUnifiedGUI(1))
    A_TrayMenu.Add()
    showAtStartup := IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1"
    A_TrayMenu.Add("Show at Startup", ToggleStartup)
    if (showAtStartup)
        A_TrayMenu.Check("Show at Startup")
    else
        A_TrayMenu.Uncheck("Show at Startup")
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", SafeExitApp)
}

ToggleTaskbar(ThisHotkey) {
    static isHidden := false
    static ABM_SETSTATE := 0xA
    static ABS_AUTOHIDE := 0x1
    static ABS_ALWAYSONTOP := 0x2
    
    size := 2 * A_PtrSize + 2 * 4 + 16 + A_PtrSize
    
    appbarData := Buffer(size, 0)
    
    NumPut("UInt", size, appbarData, 0)
    NumPut("Ptr", WinExist("ahk_class Shell_TrayWnd"), appbarData, A_PtrSize)
    
    if (isHidden) {
        NumPut("UInt", ABS_ALWAYSONTOP, appbarData, size - A_PtrSize)
        isHidden := false
        ShowNotification("➖ Taskbar Shown")
    } else {
        NumPut("UInt", ABS_AUTOHIDE, appbarData, size - A_PtrSize)
        isHidden := true
        ShowNotification("➕ Taskbar Hidden")
    }
    
    DllCall("Shell32\SHAppBarMessage", "UInt", ABM_SETSTATE, "Ptr", appbarData)
}

ToggleDesktopIcons(ThisHotkey) {
    static isHidden := false
    desktop := WinExist("ahk_class WorkerW ahk_exe explorer.exe")
    if (!desktop) {
        desktop := WinExist("ahk_class Progman ahk_exe explorer.exe")
    }
    
    if (desktop) {
        listView := ControlGetHwnd("SysListView321", "ahk_id " desktop)
        if (listView) {
            if (isHidden) {
                ControlShow("SysListView321", "ahk_id " desktop)
                isHidden := false
                ShowNotification("✅ Desktop Icons Shown")
            } else {
                ControlHide("SysListView321", "ahk_id " desktop)
                isHidden := true
                ShowNotification("🔲 Desktop Icons Hidden")
            }
        }
    }
}


ToggleMute(ThisHotkey) {
    try {
        if (SoundGetName() = "") {
            ShowNotification("❌ No audio device detected")
            return
        }
        
        SoundSetMute(-1)
        isMuted := SoundGetMute()
        ShowNotification(isMuted ? "🔇 Volume Muted" : "🔊 Volume Unmuted")
    } catch Error as err {
        ShowNotification("❌ No audio device detected")
    }
}

ToggleMic(*) {
    try {
        if (SoundGetName(, "Microphone") = "") {
            ShowNotification("❌ No microphone detected")
            return
        }
        
        SoundSetMute(-1, , "Microphone")
        isMuted := SoundGetMute(, "Microphone")
        ShowNotification(isMuted ? "🎤 Microphone Muted" : "🎤 Microphone Unmuted")
    } catch Error as err {
        ShowNotification("❌ No microphone detected")
    }
}

for action, description in hotkeyActions {
    savedHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
    if (savedHotkey != "None" && savedHotkey != "") {
        try {
            Hotkey savedHotkey, %action%
        }
    }
}

OpenExplorer(*) {
    Run("explorer.exe")
    ShowNotification("📁 Opening File Explorer")
}

OpenPowerShell(ThisHotkey) {
    downloadsPath := "C:\Users\" A_UserName "\Downloads"
    Run("powershell.exe", downloadsPath)
    ShowNotification("👾 Opening PowerShell")
}

OpenBrowser(*) {
    Run("http://")
    ShowNotification("🌐 Opening Default Browser")
}

VolumeUp(*) {
    try {
        if (SoundGetName() = "") {
            ShowNotification("❌ No audio device detected")
            return
        }
        
        SoundSetVolume("+5")
        currentVol := SoundGetVolume()
        ShowNotification("🔊 Volume: " Round(currentVol) "%")
    } catch Error as err {
        ShowNotification("❌ No audio device detected")
    }
}

VolumeDown(*) {
    try {
        if (SoundGetName() = "") {
            ShowNotification("❌ No audio device detected")
            return
        }
        
        SoundSetVolume("-5")
        currentVol := SoundGetVolume()
        ShowNotification("🔉 Volume: " Round(currentVol) "%")
    } catch Error as err {
        ShowNotification("❌ No audio device detected")
    }
}

ShowNotification(message) {
    global notificationQueue, isShowingNotification, notificationsEnabled
    
    if (!notificationsEnabled)
        return
    
    isSpecialNotification := false
    if (InStr(message, "Volume:") || InStr(message, "🔇") || InStr(message, "🎤") || 
        InStr(message, "Taskbar") || InStr(message, "Desktop Icons") ||
        InStr(message, "muted") || InStr(message, "unmuted")) {
        isSpecialNotification := true
        
        newQueue := []
        for i, item in notificationQueue {
            msg := item.text
            if !(InStr(msg, "Volume:") || InStr(msg, "🔇") || InStr(msg, "🎤") || 
                 InStr(msg, "Taskbar") || InStr(msg, "Desktop Icons") ||
                 InStr(msg, "muted") || InStr(msg, "unmuted")) {
                newQueue.Push(item)
            }
        }
        notificationQueue := newQueue
        
        for i, notifyInfo in activeNotifications {
            if (IsObject(notifyInfo) && IsObject(notifyInfo.gui)) {
                try {
                    text := notifyInfo.text
                    if (InStr(text, "Volume:") || InStr(text, "🔇") || InStr(text, "🎤") || 
                        InStr(text, "Taskbar") || InStr(text, "Desktop Icons") ||
                        InStr(text, "muted") || InStr(text, "unmuted")) {
                        notifyInfo.gui.Destroy()
                        activeNotifications.RemoveAt(i)
                    }
                }
            }
        }
    }
    
    notificationQueue.Push({text: message, isSpecial: isSpecialNotification})
    
    if (!isShowingNotification) {
        ShowNextNotification()
    }
}

ShowNextNotification() {
    global notificationQueue, isShowingNotification, currentTheme, activeNotifications, animationsEnabled
    
    if (notificationQueue.Length = 0) {
        isShowingNotification := false
        return
    }
    
    notifyInfo := notificationQueue[1]
    message := notifyInfo.text
    isSpecial := notifyInfo.isSpecial
    
    notify := Gui("-Caption +AlwaysOnTop +ToolWindow")
    notify.SetFont("s10", "Segoe UI")
    notify.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"
    
    MonitorGetWorkArea(, &left, &top, &right, &bottom)
    
    if (StrLen(message) > 40) {
        height := 50
        notify.Add("Text", "x10 y10 w280 r2 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), message)
    } else {
        notify.Add("Text", "x10 y10 w280 r1 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), message)
        height := 35
    }
    
    width := 300
    xPos := right - width - 20
    
    if (isSpecial) {
        yPos := top + 20
    } else {
        yPos := top + 20
        for i, existingNotify in activeNotifications {
            yPos += existingNotify.height + 10
        }
    }
    
    notify.Show(Format("NoActivate x{1} y{2} w{3} h{4}", xPos, yPos, width, height))
    
    ; apply fade in animation if enabled
    if (animationsEnabled) {
        FadeIn(notify.Hwnd)
    }
    
    activeNotifications.Push({gui: notify, height: height, text: message, isSpecial: isSpecial})
    
    SetTimer(RemoveNotification.Bind(notify, xPos, top), -2000)
    
    ; check again if the queue has items before removing
    if (notificationQueue.Length > 0) {
        notificationQueue.RemoveAt(1)
    }
    
    isShowingNotification := false
    
    if (notificationQueue.Length > 0) {
        SetTimer(ShowNextNotification, -10)
    }
}

RemoveNotification(notify, xPos, topPos) {
    global activeNotifications, animationsEnabled
    
    ; find and remove notification from active list
    notifyIndex := 0
    for i, notifyInfo in activeNotifications {
        if (notifyInfo.gui = notify) {
            activeNotifications.RemoveAt(i)
            notifyIndex := i
            break
        }
    }
    
    ; check if the GUI still exists before trying to fade it out
    try {
        if WinExist("ahk_id " notify.Hwnd) {
            if (animationsEnabled) {
                FadeOut(notify.Hwnd)
            }
            notify.Destroy()
        }
    } catch {
        ; GUI might already be destroyed just continue
    }
    
    ; reposition remaining notifications
    yPos := topPos + 20
    for i, notifyInfo in activeNotifications {
        try {
            if (IsObject(notifyInfo.gui) && WinExist("ahk_id " notifyInfo.gui.Hwnd)) {
                notifyInfo.gui.Move(xPos, yPos)
                yPos += notifyInfo.height + 10
            }
        } catch {
            ; skip this notification if there's an error moving it
            continue
        }
    }
}

^!m::ShowUnifiedGUI(1)



AddHotkeyRow(action, y, gui) {
    global currentTheme
    currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
    
    gui.Add("Text", "x" (gui.MarginX + 20) " y" y " w130 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), hotkeyActions[action])
    
    hotkeyText := gui.Add("Text", "x+95 w200 h25 Center Border " 
        . "Background" (currentHotkey = "None" ? "D3D3D3" : "98FB98"), 
        FormatHotkey(currentHotkey))
    
    CreateButton(gui, action, y)
}

CreateButton(keyBindWindow, action, y) {
    btn := keyBindWindow.Add("Button", "x+60 w100 h25 -E0x200", "Set Hotkey")
    btn.SetFont("c0xFFFFFF")
    btn.Opt("+Background333333")
    btn.OnEvent("Click", (*) => SetNewHotkeyGUI(action, keyBindWindow))
}

SetNewHotkeyGUI(action, parentGui) {
    global currentSetHotkeyGui, currentSetHotkeyAction, unifiedGui, isHotkeyGuiOpen
    
    if (isHotkeyGuiOpen)
        return
    
    isHotkeyGuiOpen := true
    
    currentSetHotkeyAction := action
    
    currentSetHotkeyGui := Gui("+AlwaysOnTop +MinSize200x100", "Set Hotkey")
    currentSetHotkeyGui.SetFont("s10", "Segoe UI")
    currentSetHotkeyGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"
    
    if (currentTheme = "dark") {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", currentSetHotkeyGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
    }
    
    currentSetHotkeyGui.Add("Text", "x10 y10 w300 h25 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Press the desired key combination")
    currentSetHotkeyGui.Add("Text", "x10 y35 w300 h25 c" (currentTheme = "dark" ? "0xD3D3D3" : "Gray"), "(Press Escape to clear the hotkey)")
    
    currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
    hotkeyCtrl := currentSetHotkeyGui.Add("Hotkey", "vNewHotkey w230", currentHotkey != "None" ? currentHotkey : "")
    
    helpText := currentSetHotkeyGui.Add("Link", "x+10 y72 w20 h20 -TabStop", '<a href="#">?</a>')
    helpText.SetFont("bold s10 underline c" (currentTheme = "dark" ? "0x98FB98" : "Blue"), "Segoe UI")
    helpText.OnEvent("Click", ShowPauseTooltip)
    
    ShowPauseTooltip(*) {
        tooltipText := "
        (
        Note: The Pause/Break key is supported!

        While the Hotkey input cannot display it,
        you can still use the Pause/Break key and
        it will work correctly when set.
        )"
        
        ToolTip(tooltipText, , , 1)
        SetTimer () => ToolTip(), -8000
    }
    
    currentSetHotkeyGui.OnEvent("Escape", (*) => hotkeyCtrl.Value := "")
    
    SetHotkey(inputGui, action) {
        saved := inputGui.Submit()
        newHotkey := saved.NewHotkey
        
        currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
        
        if (currentHotkey != "None") {
            try {
                Hotkey currentHotkey, "Off"
                if (activeHotkeys.Has(action)) {
                    activeHotkeys.Delete(action)
                }
            }
        }
        
        if (newHotkey = "" || newHotkey = "Escape") {
            IniWrite("None", keybindsFile, "Hotkeys", action)
            ShowNotification("⌨️ Hotkey cleared for " hotkeyActions[action])
            CleanupHotkeyGui()
            if (IsObject(unifiedGui))
                unifiedGui.Destroy()
            ShowUnifiedGUI(2)
            return
        }
        
        success := false
        try {
            try Hotkey newHotkey, "Off"
            
            fn := action
            
            Hotkey newHotkey, %fn%, "On"
            
            activeHotkeys[action] := newHotkey
            
            IniWrite(newHotkey, keybindsFile, "Hotkeys", action)
            ShowNotification("⌨️ Hotkey set to " FormatHotkey(newHotkey) " for " hotkeyActions[action])
            success := true
        }
        catch Error as err {
            ShowNotification("❌ Invalid hotkey combination: " err.Message)
        }
        
        if (success) {
            CleanupHotkeyGui()
            if (IsObject(unifiedGui))
                unifiedGui.Destroy()
            ShowUnifiedGUI(2)
        }
    }
    
    ; Function to properly clean up the hotkey GUI
    CleanupHotkeyGui(*) {
        global isHotkeyGuiOpen, currentSetHotkeyGui
        
        if (IsObject(currentSetHotkeyGui)) {
            currentSetHotkeyGui.Destroy()
            currentSetHotkeyGui := ""
        }
        
        isHotkeyGuiOpen := false
    }
    
    ; create and style buttons using the safe method
    okBtn := currentSetHotkeyGui.Add("Button", "x10 w100", "OK")
    SafeStyleButton(okBtn, currentTheme = "dark")
    okBtn.OnEvent("Click", (*) => SetHotkey(currentSetHotkeyGui, action))
    
    cancelBtn := currentSetHotkeyGui.Add("Button", "x+10 w100", "Cancel")
    SafeStyleButton(cancelBtn, currentTheme = "dark")
    cancelBtn.OnEvent("Click", CleanupHotkeyGui)
    
    ; ensure the Close event properly resets everything
    currentSetHotkeyGui.OnEvent("Close", CleanupHotkeyGui)
    
    currentSetHotkeyGui.Show("Center")
}

AddTooltip(ctl, text) {
    ctl.GetPos(&x, &y, &w, &h)
    OnMessage(0x200, WM_MOUSEMOVE)
    WM_MOUSEMOVE(wParam, lParam, msg, hwnd) {
        static PrevHwnd := 0
        if (hwnd = ctl.Hwnd && hwnd != PrevHwnd) {
            ToolTip(text)
            PrevHwnd := hwnd
        } else if (hwnd != ctl.Hwnd && PrevHwnd) {
            ToolTip()
            PrevHwnd := 0
        }
    }
}

InitHotkeys() {
    global activeHotkeys := Map()
    
    for action, description in hotkeyActions {
        hotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
        if (hotkey != "None") {
            try {
                fn := action
                Hotkey hotkey, %fn%, "On"
                activeHotkeys[action] := hotkey
            }
        }
    }
}

InitHotkeys()

Hotkey "!Backspace", ShowPowerMenu

ShowPowerMenu(*) {
    global isMenuOpen, currentTheme, powerGui, animationsEnabled
    
    if (isMenuOpen)
        return
    
    isMenuOpen := true
    
    if (IsObject(powerGui)) {
        powerGui.Destroy()
    }
    
    powerGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    powerGui.SetFont("s10", "Segoe UI")
    powerGui.BackColor := currentTheme = "dark" ? "1C1C1C" : "F0F0F0"
    
    powerGui.Add("Text", "x10 y10 w200 h30 c" (currentTheme = "dark" ? "White" : "Black") " vShutdown", "🌙  Shutdown")
    powerGui.Add("Text", "x10 y40 w200 h30 c" (currentTheme = "dark" ? "White" : "Black") " vRestart", "🔄  Restart")
    powerGui.Add("Text", "x10 y70 w200 h30 c" (currentTheme = "dark" ? "White" : "Black") " vSleep", "💤  Sleep")
    powerGui.Add("Text", "x10 y100 w200 h30 c" (currentTheme = "dark" ? "White" : "Black") " vLogoff", "🔒  Log Off")
    
    screenWidth := A_ScreenWidth
    screenHeight := A_ScreenHeight
    
    guiWidth := 220
    guiHeight := 140
    xPos := (screenWidth - guiWidth) / 2
    yPos := (screenHeight - guiHeight) / 2
    
    powerGui.Show(Format("x{1} y{2} w{3} h{4}", xPos, yPos, guiWidth, guiHeight))
    
    ; apply fade in animation if enabled
    if (animationsEnabled) {
        WinSetTransparent(0, "ahk_id " . powerGui.Hwnd)
        FadeIn(powerGui.Hwnd)
    }
    
    currentSelection := 1
    UpdateSelection(powerGui, currentSelection)
    
    Hotkey "Up", NavigateUp, "On"
    Hotkey "Down", NavigateDown, "On"
    Hotkey "Enter", ExecuteSelected, "On"
    Hotkey "Escape", ClosePowerMenu, "On"
    
    NavigateUp(*) {
        NavigateMenu("up", powerGui, &currentSelection)
    }
    
    NavigateDown(*) {
        NavigateMenu("down", powerGui, &currentSelection)
    }
    
    ExecuteSelected(*) {
        ExecutePowerAction(currentSelection)
        ClosePowerMenu()
    }
    
    ClosePowerMenu(*) {
        Hotkey "Up", "Off"
        Hotkey "Down", "Off"
        Hotkey "Enter", "Off"
        Hotkey "Escape", "Off"
        
        ; apply fade out animation if enabled
        if (animationsEnabled && IsObject(powerGui) && WinExist("ahk_id " . powerGui.Hwnd)) {
            FadeOut(powerGui.Hwnd)
        }
        
        isMenuOpen := false
        if IsObject(powerGui)
            powerGui.Destroy()
    }
    
    powerGui.OnEvent("Close", ClosePowerMenu)
}

NavigateMenu(direction, gui, &currentSelection) {
    if (direction = "up") {
        currentSelection := currentSelection > 1 ? currentSelection - 1 : 4
    } else {
        currentSelection := currentSelection < 4 ? currentSelection + 1 : 1
    }
    UpdateSelection(gui, currentSelection)
}

UpdateSelection(gui, selection) {
    if (!IsObject(gui))
        return
        
    try {
        gui["Shutdown"].SetFont("c" (currentTheme = "dark" ? "White" : "Black"))
        gui["Restart"].SetFont("c" (currentTheme = "dark" ? "White" : "Black"))
        gui["Sleep"].SetFont("c" (currentTheme = "dark" ? "White" : "Black"))
        gui["Logoff"].SetFont("c" (currentTheme = "dark" ? "White" : "Black"))
        
        highlightColor := currentTheme = "dark" ? "cLime" : "c0x008000"
        
        switch selection {
            case 1: gui["Shutdown"].SetFont(highlightColor)
            case 2: gui["Restart"].SetFont(highlightColor)
            case 3: gui["Sleep"].SetFont(highlightColor)
            case 4: gui["Logoff"].SetFont(highlightColor)
        }
    } catch Error {
        return
    }
}

ExecutePowerAction(selection) {
    switch selection {
        case 1: 
            ShowNotification("🌙 Shutting down...")
            Run("shutdown /s /t 0")
        case 2:
            ShowNotification("🔄 Restarting...")
            Run("shutdown /r /t 0")
        case 3:
            if (DllCall("powrprof\IsPwrHibernateAllowed") || DllCall("powrprof\IsPwrSuspendAllowed")) {
                ShowNotification("💤 Going to sleep...")
                DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
            } else {
                ShowNotification("❌ Sleep is not available on this system")
            }
        case 4:
            ShowNotification("🔒 Logging off...")
            Run("shutdown /l")
    }
}

NavigateUp(*) {
    global isEditorMenuOpen, currentSelection, editorGui
    if (!isEditorMenuOpen)
        return
    NavigateEditorMenu("up", editorGui, &currentSelection)
}

NavigateDown(*) {
    global isEditorMenuOpen, currentSelection, editorGui
    if (!isEditorMenuOpen)
        return
    NavigateEditorMenu("down", editorGui, &currentSelection)
}

ExecuteSelected(*) {
    global isEditorMenuOpen, currentSelection, editorGui
    if (!isEditorMenuOpen)
        return
    LaunchSelectedEditor(currentSelection, editorGui)
    CleanupEditorGui()
}

OpenVSCode(*) {
    global isEditorMenuOpen, editorGui, currentSelection, animationsEnabled
    
    if (isEditorMenuOpen) {
        CleanupEditorGui()
        return
    }
    
    hasVSCode := FindInPath("code.cmd") || FindInPath("code")
    hasVSCodium := FindInPath("codium.cmd") || FindInPath("codium")
    
    if (hasVSCode && hasVSCodium) {
        isEditorMenuOpen := true
        currentSelection := 1
        
        Send("{Alt up}{Ctrl up}{Shift up}")
        
        editorGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
        editorGui.SetFont("s10", "Segoe UI")
        editorGui.BackColor := "1C1C1C"
        
        editorGui.Add("Text", "x10 y10 w200 h30 cWhite vVSCode", "  VS Code")
        editorGui.Add("Text", "x10 y40 w200 h30 cWhite vVSCodium", "  VSCodium")
        
        screenWidth := A_ScreenWidth
        screenHeight := A_ScreenHeight
        guiWidth := 220
        guiHeight := 80
        xPos := (screenWidth - guiWidth) / 2
        yPos := (screenHeight - guiHeight) / 2
        
        editorGui.Show(Format("x{1} y{2} w{3} h{4}", xPos, yPos, guiWidth, guiHeight))
        
        ; apply fade in animation if enabled
        if (animationsEnabled) {
            WinSetTransparent(0, "ahk_id " . editorGui.Hwnd)
            FadeIn(editorGui.Hwnd)
        }
        
        UpdateEditorSelection(editorGui, currentSelection)
        
        try {
            Hotkey "Up", NavigateUp, "On"
            Hotkey "Down", NavigateDown, "On"
            Hotkey "Enter", ExecuteSelected, "On"
            Hotkey "Escape", (*) => CleanupEditorGui(), "On"
        } catch Error {
            ShowNotification("⚠️ Failed to register hotkeys for editor selection")
        }
        
        editorGui.OnEvent("Close", (*) => CleanupEditorGui())
    }
    else if (hasVSCode) {
        Run("code",, "Hide")
        ShowNotification("💻 Opening VS Code")
    }
    else if (hasVSCodium) {
        Run("codium",, "Hide")
        ShowNotification("💻 Opening VSCodium")
    }
    else {
        ShowNotification("❌ No code editor found")
    }
}

CleanupEditorGui(*) {
    global isEditorMenuOpen, editorGui, animationsEnabled
    
    try {
        Hotkey "Up", "Off"
        Hotkey "Down", "Off"
        Hotkey "Enter", "Off"
        Hotkey "Escape", "Off"
    } catch Error {
    }
    
    try {
        if IsObject(editorGui) {
            ; apply fade out animation if enabled
            if (animationsEnabled && WinExist("ahk_id " . editorGui.Hwnd)) {
                FadeOut(editorGui.Hwnd)
            }
            
            editorGui.Destroy()
            editorGui := ""
        }
    }
    
    isEditorMenuOpen := false
}

NavigateEditorMenu(direction, gui, &currentSelection) {
    if (direction = "up") {
        currentSelection := currentSelection > 1 ? currentSelection - 1 : 2
    } else {
        currentSelection := currentSelection < 2 ? currentSelection + 1 : 1
    }
    UpdateEditorSelection(gui, currentSelection)
}

UpdateEditorSelection(gui, selection) {
    gui["VSCode"].SetFont("cWhite")
    gui["VSCodium"].SetFont("cWhite")
    
    switch selection {
        case 1: gui["VSCode"].SetFont("cLime")
        case 2: gui["VSCodium"].SetFont("cLime")
    }
}

LaunchSelectedEditor(selection, gui) {
    gui.Destroy()
    switch selection {
        case 1:
            Run("code",, "Hide")
            ShowNotification("💻 Opening VS Code")
        case 2:
            Run("codium",, "Hide")
            ShowNotification("💻 Opening VSCodium")
    }
}

FindInPath(filename) {
    pathDirs := StrSplit(EnvGet("PATH"), ";")
    for dir in pathDirs {
        if (FileExist(dir "\" filename)) {
            return true
        }
    }
    return false
}

OpenCalculator(*) {
    Run("calc.exe")
    ShowNotification("🧠 Opening Calculator")
}

ResetAllHotkeys(gui) {
    gui.GetPos(&guiX, &guiY, &guiWidth)
    
    gui.Hide()
    
    result := MsgBox("Are you sure you want to reset all hotkeys?", "Reset Confirmation", "YesNo")
    
    if (result = "Yes") {
        for action, description in hotkeyActions {
            currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
            if (currentHotkey != "None") {
                try {
                    Hotkey currentHotkey, "Off"
                }
            }
        }
        
        activeHotkeys.Clear()
        
        for action, description in hotkeyActions {
            IniWrite("None", keybindsFile, "Hotkeys", action)
        }
        
        ShowNotification("🔄 All hotkeys have been reset")
        global isMenuOpen := false
        ShowUnifiedGUI(2)
    } else {
        gui.Show()
    }
}

CheckForUpdates(showMessages := false) {
    global currentVersion, versionCheckUrl, githubReleasesUrl, isCheckingForUpdates
    
    if (isCheckingForUpdates)
        return
    
    isCheckingForUpdates := true
    
    if (!InternetCheckConnection()) {
        if (showMessages) {
            MsgBox("Unable to check for updates. Please check your internet connection.", "Update Check Failed", "OK 0x30")
        }
        isCheckingForUpdates := false
        return
    }
    
    try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", versionCheckUrl, false)
        whr.Send()
        
        if (whr.Status = 200) {
            latestVersion := RegExReplace(Trim(whr.ResponseText), "[\r\n\s]")
            currentVersion := Trim(currentVersion)
           
            if (latestVersion != currentVersion) {
                result := MsgBox(
                    "A new version is available!`n`n"
                    "🍥 Current version: " currentVersion "`n"
                    "📦 Latest version: " latestVersion "`n`n"
                    "Would you like to visit the download page? (๑˃̵ᴗ˂̵)و",
                    "Update Available",
                    "YesNo 0x40"
                )
                if (result = "Yes") {
                    Run(githubReleasesUrl)
                }
            } else if (showMessages) {
                MsgBox("🍥 You have the latest version: " currentVersion, "No Updates Available ^_^", "OK 0x40")
            }
        } else if (showMessages) {
            ShowNotification("❌ Failed to check for updates (Status: " whr.Status ")")
            MsgBox("Failed to check for updates (Status: " whr.Status ")", "Update Check Failed", "OK 0x30")
        }
    } catch Error as err {
        if (InStr(err.Message, "0x80072EE7")) {
            if (showMessages) {
                MsgBox("Unable to check for updates. Please check your internet connection.", "Update Check Failed", "OK 0x30")
            }
        } else if (showMessages) {
            ShowNotification("❌ Failed to check for updates: " err.Message)
            MsgBox("Failed to check for updates: " err.Message, "Update Check Failed", "OK 0x30")
        }
    }
    
    isCheckingForUpdates := false
}

InternetCheckConnection(url := "https://www.google.com") {
    try {
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("HEAD", url, true)
        http.Send()
        http.WaitForResponse(1)
        return http.Status = 200
    } catch {
        return false
    }
}



if (IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1") {
    SetTimer((*) => ShowUnifiedGUI(1), -100)
}

SetTimer((*) => CheckForUpdates(false), -1000)

if (IniRead(launcherIniPath, "Temp", "ReopenLauncher", 0) = 1) {
    IniDelete(launcherIniPath, "Temp", "ReopenLauncher")
    SetTimer((*) => ShowUnifiedGUI(3), -100)
}

FormatHotkey(hotkeyStr) {
    if (hotkeyStr = "None" || hotkeyStr = "")
        return "None"
    
    readable := ""
    
    if (InStr(hotkeyStr, "^"))
        readable .= "Ctrl + "
    if (InStr(hotkeyStr, "!"))
        readable .= "Alt + "
    if (InStr(hotkeyStr, "+"))
        readable .= "Shift + "
    if (InStr(hotkeyStr, "#"))
        readable .= "Win + "
    
    finalKey := hotkeyStr
    finalKey := RegExReplace(finalKey, "^[!^+#]*", "")
    
    finalKey := Format("{:U}", finalKey)
    
    specialKeys := Map(
        "SC029", "Tilde",
        "SPACE", "Space",
        "TAB", "Tab",
        "ENTER", "Enter",
        "ESC", "Escape",
        "BS", "Backspace",
        "DEL", "Delete",
        "INS", "Insert",
        "HOME", "Home",
        "END", "End",
        "PGUP", "Page Up",
        "PGDN", "Page Down",
        "UP", "Up Arrow",
        "DOWN", "Down Arrow",
        "LEFT", "Left Arrow",
        "RIGHT", "Right Arrow",
        "F1", "F1",
        "F2", "F2",
        "F3", "F3",
        "F4", "F4",
        "F5", "F5",
        "F6", "F6",
        "F7", "F7",
        "F8", "F8",
        "F9", "F9",
        "F10", "F10",
        "F11", "F11",
        "F12", "F12",
        "NUMPADADD", "+",
        "NUMPADSUB", "-",
        "NUMPADMULT", "*",
        "NUMPADDIV", "/",
        "NUMPAD0", "Numpad 0",
        "NUMPAD1", "Numpad 1",
        "NUMPAD2", "Numpad 2",
        "NUMPAD3", "Numpad 3",
        "NUMPAD4", "Numpad 4",
        "NUMPAD5", "Numpad 5",
        "NUMPAD6", "Numpad 6",
        "NUMPAD7", "Numpad 7",
        "NUMPAD8", "Numpad 8",
        "NUMPAD9", "Numpad 9",
        "NUMPADDOT", "Numpad .",
        "NUMPADENTER", "Numpad Enter",
        "PRINTSCREEN", "Print Screen",
        "SCROLLLOCK", "Scroll Lock",
        "PAUSE", "Pause/Break",
        "CAPSLOCK", "Caps Lock",
        "NUMLOCK", "Num Lock",
        "MEDIA_PLAY_PAUSE", "Play/Pause",
        "MEDIA_STOP", "Media Stop",
        "MEDIA_PREV", "Previous Track",
        "MEDIA_NEXT", "Next Track",
        "VOLUME_UP", "Volume Up",
        "VOLUME_DOWN", "Volume Down",
        "VOLUME_MUTE", "Volume Mute",
        "BROWSER_BACK", "Browser Back",
        "BROWSER_FORWARD", "Browser Forward",
        "BROWSER_REFRESH", "Browser Refresh",
        "BROWSER_STOP", "Browser Stop",
        "BROWSER_SEARCH", "Browser Search",
        "BROWSER_FAVORITES", "Browser Favorites",
        "BROWSER_HOME", "Browser Home",
        "LAUNCH_MAIL", "Launch Mail",
        "LAUNCH_MEDIA", "Launch Media Player",
        "LAUNCH_APP1", "Launch App 1",
        "LAUNCH_APP2", "Launch App 2",
        "SCROLLLOCK", "Scroll Lock"
    )
    
    if (specialKeys.Has(finalKey))
        finalKey := specialKeys[finalKey]
    
    readable .= finalKey
    
    return readable
}

OpenSpotify(*) {
    if (WinExist("ahk_exe spotify.exe")) {
        WinActivate
        ShowNotification("🎵 Switching to Spotify")
        return
    }
    
    paths := [
        A_AppData "\Spotify\Spotify.exe",
        "C:\Program Files\Spotify\Spotify.exe",
        "C:\Program Files (x86)\Spotify\Spotify.exe",
        A_ProgramFiles "\Spotify\Spotify.exe",
        A_ProgramFiles . " (x86)\Spotify\Spotify.exe",
        "C:\scoop\apps\spotify\current\Spotify.exe",
        EnvGet("LOCALAPPDATA") "\Microsoft\WinGet\Packages\Spotify.Spotify_*\Spotify.exe"
    ]
    
    for path in paths {
        if (FileExist(path)) {
            Run(path)
            ShowNotification("🎵 Opening Spotify")
            return
        }
    }
    
    try {
        Run("spotify")
        ShowNotification("🎵 Opening Spotify")
    } catch {
        ShowNotification("❌ Spotify not found")
    }
}

; add function to reset all GUI states
ResetGuiStates() {
    global isHotkeyGuiOpen, isLauncherEditGuiOpen, isMenuOpen, isEditorMenuOpen
    
    isHotkeyGuiOpen := false
    isLauncherEditGuiOpen := false
    isMenuOpen := false
    isEditorMenuOpen := false
}

ToggleTheme(isDark) {
    global currentTheme, currentSetHotkeyGui, currentSetHotkeyAction, isMenuOpen, powerGui, unifiedGui, isHotkeyGuiOpen, isLauncherEditGuiOpen
    currentTheme := isDark ? "dark" : "light"
    IniWrite(currentTheme, settingsFile, "Settings", "Theme")
    
    ; Apply the theme to system tray and popup menus YaY :D
    if (isDark) {
        A_TrayMenu.Check("Dark Theme")
        darkMode.SetMode(1) ; Apply dark mode to menus
    } else {
        A_TrayMenu.Uncheck("Dark Theme")
        darkMode.SetMode(0) ; Apply light mode to menus 
    }
    
    ; Reset launcher edit GUI state if it's open
    if (isLauncherEditGuiOpen) {
        try {
            WinClose("Edit Launcher")
        } catch {
            ; Just continue if window closing fails
        }
    }
    
    if (IsObject(currentSetHotkeyGui) && WinExist("Set Hotkey")) {
        try {
            currentHotkey := currentSetHotkeyGui["NewHotkey"].Value
            currentSetHotkeyGui.Destroy()
            isHotkeyGuiOpen := false
            SetNewHotkeyGUI(currentSetHotkeyAction, "")
            currentSetHotkeyGui["NewHotkey"].Value := currentHotkey
        } catch {
            ; Continue if there's an error
        }
    }
    
    if (IsObject(powerGui)) {
        try {
            powerGui.Destroy()
        } catch {
            ; Continue if there's an error
        }
    }
    
    ; reset all GUI states to ensure clean state
    ResetGuiStates()
    
    if (IsObject(unifiedGui)) {
        currentTab := unifiedGui["TabControl"].Value
        unifiedGui.Destroy()
        ShowUnifiedGUI(currentTab)
    }
    
    ShowNotification(currentTheme = "dark" ? "🌙 Dark theme enabled" : "☀️ Light theme enabled")
}

UpdateStartupState(ctrl, *) {
    global settingsFile
    isChecked := ctrl.Value
    
    IniWrite(isChecked ? "1" : "0", settingsFile, "Settings", "ShowWelcome")
    
    if (isChecked)
        A_TrayMenu.Check("Show Welcome Message at startup")
    else
        A_TrayMenu.Uncheck("Show Welcome Message at startup")
        
    ShowNotification(isChecked ? "✅ Welcome screen enabled at startup" : "❌ Welcome screen disabled at startup")
}

global launcherIniPath := EnvGet("LOCALAPPDATA") "\WinMacros\launcher.ini"
global activeHotkeys := Map()

if !DirExist(EnvGet("LOCALAPPDATA") "\WinMacros")
    DirCreate(EnvGet("LOCALAPPDATA") "\WinMacros")


LoadLaunchers()

LoadLaunchers() {
    activeHotkeys.Clear()
    
    try {
        launcherSection := IniRead(launcherIniPath, "Launchers")
        Loop Parse, launcherSection, "`n" {
            parts := StrSplit(A_LoopField, "=")
            if parts.Length = 2 {
                hotkey := parts[1]
                path := parts[2]
                CreateLauncherHotkey(hotkey, path)
            }
        }
    }
}

LoadLaunchersToListView(lv) {
    lv.Delete()
    
    try {
        launcherSection := IniRead(launcherIniPath, "Launchers")
        nameSection := IniRead(launcherIniPath, "Names")
        
        Loop Parse, launcherSection, "`n" {
            parts := StrSplit(A_LoopField, "=")
            if parts.Length = 2 {
                hotkey := parts[1]
                path := parts[2]
                name := IniRead(launcherIniPath, "Names", hotkey, hotkey)
                readableHotkey := FormatHotkey(hotkey)
                lv.Add(, name, readableHotkey, path)
            }
        }
    }
}

CreateLauncherHotkey(hotkeyStr, path) {
    if (hotkeyStr = "" || path = "") {
        OutputDebug("Error in CreateLauncherHotkey: Empty hotkey or path")
        return false
    }
    
    if !FileExist(path) {
        OutputDebug("Error in CreateLauncherHotkey: File does not exist - " path)
        return false
    }
    
    try {
        try {
            Hotkey hotkeyStr, "Off"
        } catch {
        }
        
        name := IniRead(launcherIniPath, "Names", hotkeyStr, path)
        
        fn := (*) => (Run(path), ShowNotification("🚀 Launching " name))
        
        activeHotkeys[hotkeyStr] := fn
        
        Hotkey hotkeyStr, fn, "On"
        
        OutputDebug("Successfully registered launcher hotkey: " hotkeyStr " for " path)
        return true
    } catch Error as err {
        OutputDebug("Failed to register launcher hotkey: " hotkeyStr " - " err.Message)
        return false
    }
}

ShowUnifiedGUI(tabIndex := 1) {
    global unifiedGui, currentTabIndex, currentTheme, settingsFile
    
    currentTabIndex := tabIndex
    
    if (IsObject(unifiedGui)) {
        try {
            unifiedGui.Show()
            unifiedGui["TabControl"].Value := currentTabIndex
            return
        } catch {
        }
    }
    
    unifiedGui := Gui(, "WinMacros")
    unifiedGui.SetFont("s10", "Segoe UI")
    unifiedGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"
    
    if (currentTheme = "dark") {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", unifiedGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
    }
    
    tabs := unifiedGui.Add("Tab3", "vTabControl w770 h650", ["👋 Welcome", "⌨️ Hotkey Settings", "🚀 Launcher Settings", "🛠️ General Settings"])
    tabs.UseTab()
    
    if (currentTheme = "dark") {
        try {
            DllCall("uxtheme\SetWindowTheme", "Ptr", tabs.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
            tabs.Opt("c" . "White")
        }
    }
    
    CreateWelcomeTab(unifiedGui, tabs)
    CreateHotkeyTab(unifiedGui, tabs)
    CreateLauncherTab(unifiedGui, tabs)
    CreateSettingsTab(unifiedGui, tabs)
    
    tabs.Value := currentTabIndex
    
    unifiedGui.OnEvent("Close", (*) => unifiedGui.Hide())
    
    unifiedGui.Show("w800 h700 Center")
    
    HideFocusBorder(unifiedGui.Hwnd)
}

CreateWelcomeTab(gui, tabs) {
    global currentTheme, settingsFile, currentVersion
    
    tabs.UseTab(1)
    
    gui.SetFont("s20 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x0 y60 w770 Center", "Welcome to WinMacros!")
    
    accentColor := currentTheme = "dark" ? "0x98FB98" : "0x008000"
    gui.SetFont("s12 norm c" accentColor, "Segoe UI")
    gui.Add("Text", "x0 y100 w770 Center", "Streamline your workflow with custom Windows shortcuts")
    
    
    topBarY := 150
    topBarColor := currentTheme = "dark" ? "0x98FB98" : "0x008000"
    gui.Add("Progress", "x40 y" topBarY " w690 h3 Background" topBarColor " -Smooth Range0-100", 100)
    
    gui.SetFont("s11 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x40 y" (topBarY+20), "Helpful Shortcuts:")
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x40 y" (topBarY+50), "• Ctrl + Alt + M")
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x135 y" (topBarY+50), " - Open WinMacros")

    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x40 y" (topBarY+80), "• Alt + Backspace")
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x150 y" (topBarY+80), "- Open Power Menu")
    
    bottomBarY := topBarY + 115
    gui.Add("Progress", "x40 y" bottomBarY " w690 h3 Background" topBarColor " -Smooth Range0-100", 100)
    
    gui.SetFont("s10 c" (currentTheme = "dark" ? "0xCCCCCC" : "0x555555"), "Segoe UI")
    gui.Add("Text", "x40 y" (bottomBarY+20) " w690", "⚙️  You can customize all hotkeys and settings from the tabs above.")
    gui.Add("Text", "x40 y" (bottomBarY+45) " w690", "🔔  The application will continue running in the system tray when this window is closed.")
    gui.Add("Text", "x40 y" (bottomBarY+70) " w690", "🖱️  Right-click the tray icon for quick access to settings.")
}



CreateHotkeyTab(gui, tabs) {
    global currentTheme, hotkeyActions, keybindsFile
    
    tabs.UseTab(2)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y42 w200", "Action")
    gui.Add("Text", "x295 y42 w200", "Current Hotkey")
    gui.Add("Text", "x530 y42 w100", "Modify")
    
    y := 70
    
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "Applications")
    y += 35
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("OpenExplorer", y, gui)
    y += 35
    AddHotkeyRowToTab("OpenPowerShell", y, gui)
    y += 35
    AddHotkeyRowToTab("OpenBrowser", y, gui)
    y += 35
    AddHotkeyRowToTab("OpenVSCode", y, gui)
    y += 35
    AddHotkeyRowToTab("OpenCalculator", y, gui)
    y += 35
    AddHotkeyRowToTab("OpenSpotify", y, gui)
    
    y += 40
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "System Tools")
    y += 35
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("ToggleTaskbar", y, gui)
    y += 35
    AddHotkeyRowToTab("ToggleDesktopIcons", y, gui)
    
    y += 40
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "Sound Controls")
    y += 35
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("VolumeUp", y, gui)
    y += 35
    AddHotkeyRowToTab("VolumeDown", y, gui)
    y += 35
    AddHotkeyRowToTab("ToggleMute", y, gui)
    y += 35
    AddHotkeyRowToTab("ToggleMic", y, gui)
    
    resetBtn := gui.Add("Button", "x650 y595 w100 h30", "Reset All")
    if (currentTheme = "dark") {
        resetBtn.SetColor("0x333333", "ffffff", -1, "555555", 5)
    } else {
        resetBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 5)
    }
    resetBtn.OnEvent("Click", (*) => ResetAllHotkeysTab())
}

AddHotkeyRowToTab(action, y, gui) {
    global currentTheme, hotkeyActions, keybindsFile
    currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
    
    iconMap := Map(
        "OpenExplorer", "📁",
        "OpenPowerShell", "👾",
        "OpenBrowser", "🌐",
        "OpenVSCode", "💻",
        "OpenCalculator", "🧠",
        "OpenSpotify", "🎵",
        "ToggleTaskbar", "➖",
        "ToggleDesktopIcons", "🔲",
        "VolumeUp", "🔊",
        "VolumeDown", "🔉",
        "ToggleMute", "🔇",
        "ToggleMic", "🎤"
    )
    
    iconBgColor := currentTheme = "dark" ? "333333" : "E0E0E0"
    iconTextColor := currentTheme = "dark" ? "FFFFFF" : "000000"
    borderColor := currentTheme = "dark" ? "555555" : "CCCCCC"
    
    iconBox := gui.Add("Button", "x" (gui.MarginX + 20) " y" y " w30 h25 -E0x200", iconMap.Has(action) ? iconMap[action] : "⚙️")
    iconBox.SetFont("s11")
    
    ; Apply rounded corners and styling
    iconBox.SetBackColor("0x" iconBgColor)
    iconBox.TextColor := "0x" iconTextColor
    iconBox.BorderColor := "0x" borderColor
    iconBox.RoundedCorner := 6  ; rounded corners
    iconBox.ShowBorder := -1  ; always show border
    
    actionText := gui.Add("Text", "x+15 y" (y+2) " w145 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), hotkeyActions[action])
    
    textColor := "c0x000000"
    bgColor := currentHotkey = "None" ? "D3D3D3" : "98FB98"
    hotkeyText := gui.Add("Text", "x+20 y" y " w200 h22 Center Border " textColor " Background" bgColor,
        FormatHotkey(currentHotkey))
    
    CreateButtonInTab(gui, action, y)
}

CreateButtonInTab(gui, action, y) {
    global currentTheme
    
    btn := gui.Add("Button", "x+60 y" y " w100 h25 -E0x200", "Set Hotkey")
    if (currentTheme = "dark") {
        btn.SetColor("0x333333", "ffffff", -1, "555555", 9)
    } else {
        btn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
    }
    btn.OnEvent("Click", (*) => SetNewHotkeyGUI(action, unifiedGui))
}

ResetAllHotkeysTab() {
    global unifiedGui
    
    result := MsgBox("Are you sure you want to reset all hotkeys?", "Reset Confirmation", "YesNo")
    
    if (result = "Yes") {
        for action, description in hotkeyActions {
            currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
            if (currentHotkey != "None") {
                try {
                    Hotkey currentHotkey, "Off"
                }
            }
        }
        
        activeHotkeys.Clear()
        
        for action, description in hotkeyActions {
            IniWrite("None", keybindsFile, "Hotkeys", action)
        }
        
        ShowNotification("🔄 All hotkeys have been reset")
        
        
        if (IsObject(unifiedGui)) {
                unifiedGui.Destroy()
                ShowUnifiedGUI(2)
            }
    }
}

CreateLauncherTab(gui, tabs) {
    global currentTheme, launcherIniPath
    
    tabs.UseTab(3)
    
    lv := gui.Add("ListView", "x20 y40 w730 h320 Grid -Multi NoSortHdr +LV0x10000 +HScroll +VScroll", ["Name", "Hotkey", "Path"])
    
    if (currentTheme = "dark") {
        lv.Opt("+Background333333 cWhite")
        DllCall("uxtheme\SetWindowTheme", "Ptr", lv.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    }
    
    totalWidth := 730
    nameWidth := 150
    hotkeyWidth := 150
    pathWidth := totalWidth - nameWidth - hotkeyWidth - 4
    
    lv.ModifyCol(1, nameWidth)
    lv.ModifyCol(2, hotkeyWidth)
    lv.ModifyCol(3, pathWidth)
    
    LoadLaunchersToListViewTab(lv)
    
    gui.Add("Text", "x20 y380 w100 c" (currentTheme = "dark" ? "White" : "Black"), "🏷️ Name:")
    nameInput := gui.Add("Edit", "x120 y380 w200 h23 vLauncherName " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    
    gui.Add("Text", "x20 y410 w100 c" (currentTheme = "dark" ? "White" : "Black"), "⌨️ Hotkey:")
    hotkeyInput := gui.Add("Hotkey", "x120 y410 w200 vLauncherHotkey")
    helpText := gui.Add("Link", "x+10 y413 w20 h20 -TabStop", '<a href="#">?</a>')
    helpText.SetFont("bold s10 underline c" (currentTheme = "dark" ? "0x98FB98" : "Blue"), "Segoe UI")
    helpText.OnEvent("Click", ShowLauncherHotkeyTooltip)
    
    gui.OnEvent("Escape", (*) => hotkeyInput.Value := "")

    gui.Add("Text", "x20 y440 w100 c" (currentTheme = "dark" ? "White" : "Black"), "📂 Path:")
    pathInput := gui.Add("Edit", "x120 y440 w500 h23 vLauncherPath " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    browseBtn := gui.Add("Button", "x630 y438 w100", "Browse")
    addBtn := gui.Add("Button", "x120 y480 w100", "Add")
    editBtn := gui.Add("Button", "x230 y480 w100", "Edit")
    deleteBtn := gui.Add("Button", "x340 y480 w100", "Delete")
    
    if (currentTheme = "dark") {
        browseBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        addBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        editBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        deleteBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
    } else {
        browseBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        addBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        editBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        deleteBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
    }
    
    browseBtn.OnEvent("Click", BrowseLauncherFile)
    addBtn.OnEvent("Click", AddLauncherFromTab)
    editBtn.OnEvent("Click", EditLauncherFromTab)
    deleteBtn.OnEvent("Click", DeleteLauncherFromTab)
}

ShowLauncherHotkeyTooltip(*) {
    tooltipText := "
    (
    Note: The Pause/Break key is supported!

    While the Hotkey input cannot display it,
    you can still use the Pause/Break key and
    it will work correctly when set.
    )"
    
    ToolTip(tooltipText, , , 1)
    SetTimer () => ToolTip(), -8000
}

BrowseLauncherFile(*) {
    global unifiedGui
    
    if (selected := FileSelect(3,, "Select Program", "Programs (*.exe)")) {
        unifiedGui["LauncherPath"].Value := selected
    }
}

LoadLaunchersToListViewTab(lv) {
    global launcherIniPath
    
    lv.Delete()
    
    try {
        launcherSection := IniRead(launcherIniPath, "Launchers")
        nameSection := IniRead(launcherIniPath, "Names")
        
        Loop Parse, launcherSection, "`n" {
            parts := StrSplit(A_LoopField, "=")
            if parts.Length = 2 {
                hotkey := parts[1]
                path := parts[2]
                name := IniRead(launcherIniPath, "Names", hotkey, hotkey)
                readableHotkey := FormatHotkey(hotkey)
                lv.Add(, name, readableHotkey, path)
            }
        }
    }
}

AddLauncherFromTab(*) {
    global unifiedGui, launcherIniPath
    
    nameInput := unifiedGui["LauncherName"].Value
    hotkeyInput := unifiedGui["LauncherHotkey"].Value
    pathInput := unifiedGui["LauncherPath"].Value
    
    if (nameInput = "" || hotkeyInput = "" || pathInput = "") {
        ShowNotification("❌ Please fill in all fields")
        return
    }
    
    try {
        launcherSection := IniRead(launcherIniPath, "Launchers")
        nameSection := IniRead(launcherIniPath, "Names")
        
        Loop Parse, launcherSection, "`n" {
            parts := StrSplit(A_LoopField, "=")
            if parts.Length = 2 {
                existingPath := parts[2]
                existingName := IniRead(launcherIniPath, "Names", parts[1], "")
                
                if (existingPath = pathInput) {
                    ShowNotification("❌ A launcher with this path already exists")
                    return
                }
                if (existingName = nameInput) {
                    ShowNotification("❌ A launcher with this name already exists")
                    return
                }
            }
        }
    }
    
    IniWrite(pathInput, launcherIniPath, "Launchers", hotkeyInput)
    IniWrite(nameInput, launcherIniPath, "Names", hotkeyInput)
    
    CreateLauncherHotkey(hotkeyInput, pathInput)
    
    unifiedGui["LauncherName"].Value := ""
    unifiedGui["LauncherHotkey"].Value := ""
    unifiedGui["LauncherPath"].Value := ""
    
    try {
        lv := unifiedGui["SysListView321"]
        LoadLaunchersToListViewTab(lv)
    } catch {
        currentTabIndex := 3
        ShowUnifiedGUI(currentTabIndex)
    }
    
    ShowNotification("✅ Launcher added successfully")
}

EditLauncherFromTab(*) {
    global unifiedGui, launcherIniPath, currentTheme, isLauncherEditGuiOpen
    
    if (isLauncherEditGuiOpen)
        return
        
    try {
        lv := unifiedGui["SysListView321"]
        if !(rowNum := lv.GetNext()) {
            ShowNotification("❌ Please select a launcher to edit")
            return
        }
        
        isLauncherEditGuiOpen := true
        
        oldName := lv.GetText(rowNum, 1)
        readableHotkey := lv.GetText(rowNum, 2)
        oldPath := lv.GetText(rowNum, 3)
        
        hotkeyParts := StrSplit(readableHotkey, " + ")
        oldHotkey := ""
        for part in hotkeyParts {
            switch part {
                case "Ctrl": oldHotkey .= "^"
                case "Alt": oldHotkey .= "!"
                case "Shift": oldHotkey .= "+"
                case "Win": oldHotkey .= "#"
                case "Caps Lock": oldHotkey := "CapsLock"
                case "Pause/Break": oldHotkey := "Pause"
                case "Scroll Lock": oldHotkey := "ScrollLock"
                case "Page Up": oldHotkey := "PgUp"
                case "Page Down": oldHotkey := "PgDn"
                default: oldHotkey .= Format("{:L}", part)
            }
        }
        
        editGui := Gui("+AlwaysOnTop +MinSize640x150", "Edit Launcher")
        editGui.SetFont("s10", "Segoe UI")
        editGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "FFFFFF"
        
        if (currentTheme = "dark") {
            DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", editGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
            DllCall("uxtheme\SetWindowTheme", "Ptr", editGui.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
        }
        
        editGui.Add("Text", "x10 y10 w100 c" (currentTheme = "dark" ? "White" : "Black"), "🏷️ Name:")
        editNameInput := editGui.Add("Edit", "x110 y10 w200 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), oldName)
        
        editGui.Add("Text", "x10 y40 w100 c" (currentTheme = "dark" ? "White" : "Black"), "⌨️ Hotkey:")
        editHotkeyInput := editGui.Add("Hotkey", "x110 y40 w200", oldHotkey)
        
        editGui.OnEvent("Escape", (*) => editHotkeyInput.Value := "")

        editGui.Add("Text", "x10 y70 w100 c" (currentTheme = "dark" ? "White" : "Black"), "📂 Path:")
        editPathInput := editGui.Add("Edit", "x110 y70 w400 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), oldPath)
        editBrowseBtn := editGui.Add("Button", "x520 y68 w100", "Browse")
        
        saveBtn := editGui.Add("Button", "x420 y110 w100", "Save")
        cancelBtn := editGui.Add("Button", "x530 y110 w100", "Cancel")
        
        ; Style buttons using the safe method
        SafeStyleButton(editBrowseBtn, currentTheme = "dark")
        SafeStyleButton(saveBtn, currentTheme = "dark")
        SafeStyleButton(cancelBtn, currentTheme = "dark")
        
        BrowsePath(*) {
            if (selected := FileSelect(3,, "Select Program", "Programs (*.exe)"))
                editPathInput.Value := selected
        }
        
        SaveChanges(*) {
            if (editNameInput.Value = "" || editHotkeyInput.Value = "" || editPathInput.Value = "") {
                ShowNotification("❌ Please fill in all fields")
                return
            }
            
            try {
                Hotkey oldHotkey, "Off"
                activeHotkeys.Delete(oldHotkey)
            }
            
            try {
                launcherSection := IniRead(launcherIniPath, "Launchers")
                Loop Parse, launcherSection, "`n" {
                    parts := StrSplit(A_LoopField, "=")
                    if parts.Length = 2 {
                        existingHotkey := parts[1]
                        existingPath := parts[2]
                        existingName := IniRead(launcherIniPath, "Names", existingHotkey, "")
                        
                        if (existingPath = oldPath && existingName = oldName)
                            continue
                            
                        if (existingPath = editPathInput.Value) {
                            ShowNotification("❌ A launcher with this path already exists")
                            return
                        }
                        if (existingName = editNameInput.Value) {
                            ShowNotification("❌ A launcher with this name already exists")
                            return
                        }
                        if (existingHotkey = editHotkeyInput.Value) {
                            ShowNotification("❌ This hotkey is already in use")
                            return
                        }
                    }
                }
            }
            
            newHotkey := editHotkeyInput.Value
            newPath := editPathInput.Value
            newName := editNameInput.Value
            
            IniDelete(launcherIniPath, "Launchers", oldHotkey)
            IniDelete(launcherIniPath, "Names", oldHotkey)
            
            IniWrite(newPath, launcherIniPath, "Launchers", newHotkey)
            IniWrite(newName, launcherIniPath, "Names", newHotkey)
            
            CleanupEditLauncherGui()
            
            ShowNotification("✅ Launcher updated successfully")
            
            try {
                lv := unifiedGui["SysListView321"]
                LoadLaunchersToListViewTab(lv)
            } catch {
                currentTabIndex := 3
                ShowUnifiedGUI(currentTabIndex)
            }
            
            CreateLauncherHotkey(newHotkey, newPath)
        }
        
        editBrowseBtn.OnEvent("Click", BrowsePath)
        saveBtn.OnEvent("Click", SaveChanges)
        
        ; Make sure to reset the isLauncherEditGuiOpen flag when closing
        CleanupEditLauncherGui(*) {
            global isLauncherEditGuiOpen
            
            if (IsObject(editGui)) {
                editGui.Destroy()
            }
            
            isLauncherEditGuiOpen := false
        }
        
        cancelBtn.OnEvent("Click", CleanupEditLauncherGui)
        
        ; Ensure the Close event properly resets everything
        editGui.OnEvent("Close", CleanupEditLauncherGui)
        
        ; Center the window on screen
        editGui.Show("w640 h150 Center")
    } catch Error as err {
        isLauncherEditGuiOpen := false
        ShowNotification("❌ Error: " err.Message)
    }
}

DeleteLauncherFromTab(*) {
    global unifiedGui, launcherIniPath
    
    try {
        lv := unifiedGui["SysListView321"]
        if !(rowNum := lv.GetNext()) {
            ShowNotification("❌ Please select a launcher to delete")
            return
        }
        
        name := lv.GetText(rowNum, 1)
        readableHotkey := lv.GetText(rowNum, 2)
        
        hotkeyParts := StrSplit(readableHotkey, " + ")
        ahkHotkey := ""
        for part in hotkeyParts {
            switch part {
                case "Ctrl": ahkHotkey .= "^"
                case "Alt": ahkHotkey .= "!"
                case "Shift": ahkHotkey .= "+"
                case "Win": ahkHotkey .= "#"
                case "Caps Lock": ahkHotkey := "CapsLock"
                case "Pause/Break": ahkHotkey := "Pause"
                case "Scroll Lock": ahkHotkey := "ScrollLock"
                case "Page Up": ahkHotkey := "PgUp"
                case "Page Down": ahkHotkey := "PgDn"
                default: ahkHotkey .= Format("{:L}", part)
            }
        }
        
        try {
            Hotkey ahkHotkey, "Off"
            activeHotkeys.Delete(ahkHotkey)
        }
        
        IniDelete(launcherIniPath, "Launchers", ahkHotkey)
        IniDelete(launcherIniPath, "Names", ahkHotkey)
        
        LoadLaunchersToListViewTab(lv)
        
        ShowNotification("🗑️ Deleted launcher: " name)
    } catch Error as err {
        ShowNotification("❌ Error: " err.Message)
    }
}

CreateSettingsTab(gui, tabs) {
    global currentTheme, settingsFile, notificationsEnabled, currentVersion, animationsEnabled
    
    tabs.UseTab(4)
    
    gui.SetFont("s12 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x20 y40 w730", "General Settings")
    
    gui.Add("Text", "x20 y70 w730 h2 0x10")
    
    leftColX := 40
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y90 w200", "Appearance")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    themeToggle := gui.Add("Checkbox", "x" (leftColX+20) " y120 vThemeToggle", "Dark Theme")
    themeToggle.Value := currentTheme = "dark"
    themeToggle.OnEvent("Click", (*) => ToggleTheme(themeToggle.Value))
    
    animChk := gui.Add("Checkbox", "x" (leftColX+20) " y150 vEnableAnimations", "Enable animations")
    animChk.Value := animationsEnabled
    animChk.OnEvent("Click", ToggleAnimationsFromSettings)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y190 w200", "Startup Options")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    showAtStartup := IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1"
    welcomeChk := gui.Add("Checkbox", "x" (leftColX+20) " y220 vShowWelcome", "Show welcome screen at startup")
    welcomeChk.Value := showAtStartup
    welcomeChk.OnEvent("Click", ToggleWelcomeStartup)
    
    isStartupEnabled := FileExist(A_Startup "\WinMacros.lnk") ? true : false
    startupChk := gui.Add("Checkbox", "x" (leftColX+20) " y250 vRunOnStartup", "Run WinMacros on Windows startup")
    startupChk.Value := isStartupEnabled
    startupChk.OnEvent("Click", ToggleWindowsStartupFromSettings)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y290 w200", "Notifications")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    notifyChk := gui.Add("Checkbox", "x" (leftColX+20) " y320 vEnableNotifications", "Enable notifications")
    notifyChk.Value := notificationsEnabled
    notifyChk.OnEvent("Click", ToggleNotificationsFromSettings)
    
    rightColX := 400
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" rightColX " y90 w200", "Configuration")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" (rightColX+20) " y120", "Hotkeys:")
    
    exportHotkeysBtn := gui.Add("Button", "x" (rightColX+100) " y118 w100 h25", "Export")
    importHotkeysBtn := gui.Add("Button", "x" (rightColX+210) " y118 w100 h25", "Import")
    
    gui.Add("Text", "x" (rightColX+20) " y155", "Launchers:")
    
    exportLaunchersBtn := gui.Add("Button", "x" (rightColX+100) " y153 w100 h25", "Export")
    importLaunchersBtn := gui.Add("Button", "x" (rightColX+210) " y153 w100 h25", "Import")
    
    openConfigBtn := gui.Add("Button", "x" (rightColX+20) " y190 w290 h25", "📁 Open Config Location")
    
    if (currentTheme = "dark") {
        openConfigBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        exportHotkeysBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        importHotkeysBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        exportLaunchersBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
        importLaunchersBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
    } else {
        openConfigBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        exportHotkeysBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        importHotkeysBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        importLaunchersBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
        exportLaunchersBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
    }
    
    exportHotkeysBtn.OnEvent("Click", ExportHotkeys)
    importHotkeysBtn.OnEvent("Click", ImportHotkeys)
    exportLaunchersBtn.OnEvent("Click", ExportLaunchers)
    importLaunchersBtn.OnEvent("Click", ImportLaunchers)
    openConfigBtn.OnEvent("Click", OpenConfigLocation)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" rightColX " y230 w200", "Updates")
    
    checkUpdateBtn := gui.Add("Button", "x" (rightColX+20) " y260 w200 h30", "Check for Updates")
    if (currentTheme = "dark") {
        checkUpdateBtn.SetColor("0x333333", "ffffff", -1, "555555", 9)
    } else {
        checkUpdateBtn.SetColor("ffffff", "0x333333", -1, "CCCCCC", 9)
    }
    checkUpdateBtn.OnEvent("Click", (*) => CheckForUpdates(true))
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y360 w200", "About")
    
    gui.SetFont("s10 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" (leftColX+20) " y390", "[#] Current Version: " currentVersion)
    
    gui.SetFont("s10", "Segoe UI")
    linkColor := currentTheme = "dark" ? "cWhite" : "cBlue"
    githubLink := gui.Add("Link", "x" (leftColX+20) " y420 w300 -TabStop " linkColor, 
        '<a href="https://github.com/fr0st-iwnl/WinMacros">GitHub Repository</a> | <a href="https://winmacros.netlify.app/">Website</a>')
}

ToggleWelcomeStartup(ctrl, *) {
    global settingsFile
    isChecked := ctrl.Value
    
    IniWrite(isChecked ? "1" : "0", settingsFile, "Settings", "ShowWelcome")
    
    if (isChecked) {
        A_TrayMenu.Check("Show Welcome Message at startup")
        ShowNotification("✅ Welcome screen enabled at startup")
    } else {
        A_TrayMenu.Uncheck("Show Welcome Message at startup")
        ShowNotification("❌ Welcome screen disabled at startup")
    }
}

ToggleWindowsStartupFromSettings(ctrl, *) {
    global unifiedGui
    startupPath := A_Startup "\WinMacros.lnk"
    
    if (!FileExist(startupPath) && ctrl.Value) {
        try {
            FileCreateShortcut(A_ScriptFullPath, startupPath,, "Launch WinMacros on startup")
            ShowNotification("✅ WinMacros will run on Windows startup")
            A_TrayMenu.Check("Run on Windows Startup")
        } catch Error as err {
            ShowNotification("❌ Failed to create startup shortcut")
            ctrl.Value := false
        }
    } else if (FileExist(startupPath) && !ctrl.Value) {
        try {
            FileDelete(startupPath)
            ShowNotification("❌ WinMacros will not run on Windows startup")
            A_TrayMenu.Uncheck("Run on Windows Startup")
        } catch Error as err {
            ShowNotification("❌ Failed to remove startup shortcut")
            ctrl.Value := true
        }
    }
}

ToggleNotificationsFromSettings(ctrl, *) {
    global notificationsEnabled, settingsFile
    
    notificationsEnabled := ctrl.Value
    IniWrite(notificationsEnabled ? "1" : "0", settingsFile, "Settings", "Notifications")
    
    if (notificationsEnabled) {
        A_TrayMenu.Check("Enable Notifications")
        ShowNotification("📢 Notifications enabled")
    } else {
        A_TrayMenu.Uncheck("Enable Notifications")
    }
}

ToggleAnimationsFromSettings(ctrl, *) {
    global animationsEnabled, settingsFile
    
    animationsEnabled := ctrl.Value
    IniWrite(animationsEnabled ? "1" : "0", settingsFile, "Settings", "Animations")
    
    if (animationsEnabled) {
        A_TrayMenu.Check("Enable Animations")
        ShowNotification("✅ Animations enabled")
    } else {
        A_TrayMenu.Uncheck("Enable Animations")
        ShowNotification("❌ Animations disabled")
    }
}

ToggleAnimations(*) {
    global animationsEnabled, settingsFile, unifiedGui
    
    animationsEnabled := !animationsEnabled
    IniWrite(animationsEnabled ? "1" : "0", settingsFile, "Settings", "Animations")
    
    if (animationsEnabled) {
        A_TrayMenu.Check("Enable Animations")
        ShowNotification("✅ Animations enabled")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["EnableAnimations"]) {
                    unifiedGui["EnableAnimations"].Value := true
                }
            }
        }
    } else {
        A_TrayMenu.Uncheck("Enable Animations")
        ShowNotification("❌ Animations disabled")
        
        if (IsObject(unifiedGui) && WinExist("WinMacros")) {
            try {
                if (unifiedGui["EnableAnimations"]) {
                    unifiedGui["EnableAnimations"].Value := false
                }
            }
        }
    }
}

try {
    showAfterReload := IniRead(settingsFile, "Temp", "ShowAfterReload", "0")
    if (showAfterReload = "1") {
        lastTabIndex := Integer(IniRead(settingsFile, "Temp", "LastTabIndex", "1"))
        IniDelete(settingsFile, "Temp", "ShowAfterReload")
        IniDelete(settingsFile, "Temp", "LastTabIndex")
        
        SetTimer(() => ShowUnifiedGUI(lastTabIndex), -500)
        
        if (lastTabIndex = 2) {
            SetTimer(() => ShowNotification("✅ Hotkeys activated successfully"), -1000)
        } else if (lastTabIndex = 3) {
            SetTimer(() => ShowNotification("✅ Launchers activated successfully"), -1000)
        }
    }
} catch {
}


; Import/Export Functions [Hotkeys & Launchers] (hated writing this shitty code)
ExportHotkeys(*) {
    global keybindsFile
    
    if (filePath := FileSelect("S", A_Desktop "\WinMacros_Hotkeys.ini", "Export Hotkeys", "INI Files (*.ini)")) {
        try {
            if (!FileExist(keybindsFile)) {
                for action, description in hotkeyActions {
                    IniWrite("None", keybindsFile, "Hotkeys", action)
                }
                ShowNotification("✅ Created new hotkeys file and exported it")
            } else {
                FileCopy(keybindsFile, filePath, true)
                ShowNotification("✅ Hotkeys exported successfully")
            }
        } catch Error as err {
            ShowNotification("❌ Failed to export hotkeys: " err.Message)
        }
    }
}

ImportHotkeys(*) {
    global keybindsFile, hotkeyActions, activeHotkeys, unifiedGui, currentTabIndex, settingsFile
    
    if (filePath := FileSelect(3,, "Import Hotkeys", "INI Files (*.ini)")) {
        try {
            if !DirExist(EnvGet("LOCALAPPDATA") "\WinMacros") {
                DirCreate(EnvGet("LOCALAPPDATA") "\WinMacros")
            }
            
            if (FileExist(keybindsFile)) {
                try {
                    FileCopy(keybindsFile, keybindsFile ".backup", true)
                } catch Error as backupErr {
                    OutputDebug("Backup failed: " backupErr.Message)
                }
                
                for action, description in hotkeyActions {
                    currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
                    if (currentHotkey != "None") {
                        try {
                            Hotkey currentHotkey, "Off"
                        } catch {
                        }
                    }
                }
            }
            
            activeHotkeys.Clear()
            
            try {
                FileCopy(filePath, keybindsFile, true)
            } catch Error as copyErr {
                if (InStr(copyErr.Message, "destination path does not exist")) {
                    SplitPath(keybindsFile, , &dir)
                    if (dir && !DirExist(dir)) {
                        DirCreate(dir)
                        FileCopy(filePath, keybindsFile, true)
                    } else {
                        throw copyErr
                    }
                } else {
                    throw copyErr
                }
            }
            
            for action, description in hotkeyActions {
                newHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
                if (newHotkey != "None") {
                    try {
                        fn := action
                        Hotkey newHotkey, %fn%, "On"
                        activeHotkeys[action] := newHotkey
                    } catch Error as err {
                        ShowNotification("⚠️ Warning: Could not set hotkey for " description)
                    }
                }
            }
            
            ShowNotification("✅ Hotkeys imported successfully")
            
            IniWrite(2, settingsFile, "Temp", "LastTabIndex")
            IniWrite(1, settingsFile, "Temp", "ShowAfterReload")
            
            Reload()
            
            if (IsObject(unifiedGui)) {
                unifiedGui.Destroy()
                ShowUnifiedGUI(2)
            }
        } catch Error as err {
            ShowNotification("❌ Failed to import hotkeys: " err.Message)
        }
    }
}

ExportLaunchers(*) {
    global launcherIniPath
    
    if (filePath := FileSelect("S", A_Desktop "\WinMacros_Launchers.ini", "Export Launchers", "INI Files (*.ini)")) {
        try {
            if (!FileExist(launcherIniPath)) {
                FileAppend("", launcherIniPath)
                ShowNotification("✅ Created empty launcher file and exported it")
            }
            FileCopy(launcherIniPath, filePath, true)
            ShowNotification("✅ Launchers exported successfully")
        } catch Error as err {
            ShowNotification("❌ Failed to export launchers: " err.Message)
        }
    }
}

ImportLaunchers(*) {
    global launcherIniPath, activeHotkeys, unifiedGui, currentTabIndex, settingsFile
    
    if (filePath := FileSelect(3,, "Import Launchers", "INI Files (*.ini)")) {
        try {
            if !DirExist(EnvGet("LOCALAPPDATA") "\WinMacros") {
                DirCreate(EnvGet("LOCALAPPDATA") "\WinMacros")
            }
            
            if (FileExist(launcherIniPath)) {
                try {
                    FileCopy(launcherIniPath, launcherIniPath ".backup", true)
                } catch Error as backupErr {
                    OutputDebug("Backup failed: " backupErr.Message)
                }
                
                try {
                    launcherSection := IniRead(launcherIniPath, "Launchers")
                    Loop Parse, launcherSection, "`n" {
                        parts := StrSplit(A_LoopField, "=")
                        if parts.Length = 2 {
                            try {
                                Hotkey parts[1], "Off"
                            } catch {
                            }
                        }
                    }
                } catch {
                }
                
                for key, value in activeHotkeys.Clone() {
                    if (IsObject(value)) {
                        activeHotkeys.Delete(key)
                    }
                }
            }
            
            try {
                FileCopy(filePath, launcherIniPath, true)
            } catch Error as copyErr {
                if (InStr(copyErr.Message, "destination path does not exist")) {
                    SplitPath(launcherIniPath, , &dir)
                    if (dir && !DirExist(dir)) {
                        DirCreate(dir)
                        FileCopy(filePath, launcherIniPath, true)
                    } else {
                        throw copyErr
                    }
                } else {
                    throw copyErr
                }
            }
            
            IniWrite(3, settingsFile, "Temp", "LastTabIndex")
            IniWrite(1, settingsFile, "Temp", "ShowAfterReload")
            
            ShowNotification("✅ Launchers imported - Activating...")
            Sleep(800)
            
            Reload()
            
            if (IsObject(unifiedGui)) {
                unifiedGui.Destroy()
                ShowUnifiedGUI(3)
            }
        } catch Error as err {
            ShowNotification("❌ Failed to import launchers: " err.Message)
        }
    }
}

OpenConfigLocation(*) {
    Run("explorer.exe " EnvGet("LOCALAPPDATA") "\WinMacros")
}

;                                       THE END OF THE SCRIPT
; ------------------------------------------------------------------------------------------------

; THIS MAKES IT FASTER TO TEST THE SCRIPT
; ^!r::TestingScript() 

; TESTING SCRIPT FUNCTION TO RELOAD SCRIPT AND GO TO A SPECIFIC TAB
TestingScript(*) {
    global settingsFile
    IniWrite(1, settingsFile, "Temp", "LastTabIndex")  ; Change number for different tab
    IniWrite(1, settingsFile, "Temp", "ShowAfterReload")
    Reload()
}


; ╔════════════════════════════════════════════════════════════════════════════════════════════════════════╗
; ║                                          LIBRARIES                                                     ║
; ╠════════════════════════════════════════════════════════════════════════════════════════════════════════╣
; ║                                                                                                        ║
; ║  ; #Include <ColorButton>                ; https://github.com/nperovic/ColorButton.ahk                 ║
; ║  ; #Include <DarkMsgBox>                 ; https://github.com/nperovic/DarkMsgBox                      ║
; ║  ; #Include <SystemThemeAwareToolTip>    ; https://github.com/nperovic/SystemThemeAwareToolTip         ║
; ║                                                                                                        ║
; ╠════════════════════════════════════════════════════════════════════════════════════════════════════════╣
; ║                                      SPECIAL THANKS TO:                                                ║
; ║                                                                                                        ║
; ║  - nperovic | https://github.com/nperovic                                                              ║
; ║  - AutoHotkey community (for awesome scripts)                                                          ║
; ╚════════════════════════════════════════════════════════════════════════════════════════════════════════╝



; ========================================= COLOR BUTTON =========================================

; StructFromPtr(StructClass, Address) => StructClass(Address)

Buffer.Prototype.PropDesc := PropDesc
PropDesc(buf, name, ofst, type, ptr?) {
    if (ptr??0)
        NumPut(type, NumGet(ptr, ofst, type), buf, ofst)
    buf.DefineProp(name, {
        Get: NumGet.Bind(, ofst, type),
        Set: (p, v) => NumPut(type, v, buf, ofst)
    })
}

class NMHDR extends Buffer {
    __New(ptr?) {
        super.__New(A_PtrSize * 2 + 4)
        this.PropDesc("hwndFrom", 0, "uptr", ptr?)
        this.PropDesc("idFrom", A_PtrSize,"uptr", ptr?)   
        this.PropDesc("code", A_PtrSize * 2 ,"int", ptr?)     
    }
}

class RECT extends Buffer { 
    __New(ptr?) {
        super.__New(16, 0)
        for i, prop in ["left", "top", "right", "bottom"]
            this.PropDesc(prop, 4 * (i-1), "int", ptr?)
        this.DefineProp("Width", {Get: rc => (rc.right - rc.left)})
        this.DefineProp("Height", {Get: rc => (rc.bottom - rc.top)})
    }
}

class NMCUSTOMDRAWINFO extends Buffer
{
    __New(ptr?) {
        static x64 := (A_PtrSize = 8)
        super.__New(x64 ? 80 : 48)
        this.hdr := NMHDR(ptr?)
        this.rc  := RECT((ptr??0) ? ptr + (x64 ? 40 : 20) : unset)
        this.PropDesc("dwDrawStage", x64 ? 24 : 12, "uint", ptr?)  
        this.PropDesc("hdc"        , x64 ? 32 : 16, "uptr", ptr?)          
        this.PropDesc("dwItemSpec" , x64 ? 56 : 36, "uptr", ptr?)   
        this.PropDesc("uItemState" , x64 ? 64 : 40, "int", ptr?)   
        this.PropDesc("lItemlParam", x64 ? 72 : 44, "iptr", ptr?)
    }
}

; ================================================================================ 

/**
 * The extended class for the built-in `Gui.Button` class.
 * @method SetBackColor Set the button's background color
 * @example
 * btn := myGui.AddButton(, "SUPREME")
 * btn.SetBackColor(0xaa2031)
 */
class _BtnColor extends Gui.Button
{
    static __New() {
        for prop in this.Prototype.OwnProps()
            if (!super.Prototype.HasProp(prop) && SubStr(prop, 1, 1) != "_")
                super.Prototype.DefineProp(prop, this.Prototype.GetOwnPropDesc(prop))

        Gui.CheckBox.Prototype.DefineProp("GetTextFlags", this.Prototype.GetOwnPropDesc("GetTextFlags"))
        Gui.Radio.Prototype.DefineProp("GetTextFlags", this.Prototype.GetOwnPropDesc("GetTextFlags"))
    }

    GetTextFlags(&center?, &vcenter?, &right?, &bottom?)
    {
        static BS_BOTTOM     := 0x800
        static BS_CENTER     := 0x300
        static BS_LEFT       := 0x100
        static BS_LEFTTEXT   := 0x20
        static BS_MULTILINE  := 0x2000
        static BS_RIGHT      := 0x200
        static BS_TOP        := 0x0400
        static BS_VCENTER    := 0x0C00
        static DT_BOTTOM     := 0x8
        static DT_CENTER     := 0x1
        static DT_LEFT       := 0x0
        static DT_RIGHT      := 0x2
        static DT_SINGLELINE := 0x20
        static DT_TOP        := 0x0
        static DT_VCENTER    := 0x4
        static DT_WORDBREAK  := 0x10
        
        dwStyle     := ControlGetStyle(this)
        txC         := dwStyle & BS_CENTER
        txR         := dwStyle & BS_RIGHT
        txL         := dwStyle & BS_LEFT
        dwTextFlags := (dwStyle & BS_BOTTOM) ? DT_BOTTOM : !(dwStyle & BS_TOP) ? DT_VCENTER : DT_TOP
        
        if (this.Type = "Button") 
            dwTextFlags |= (txC && txR && !txL) ? DT_RIGHT : (txC && txL && !txR) ? DT_LEFT : DT_CENTER
        else
            dwTextFlags |= txL && txR ? DT_CENTER : !txL && txR ? DT_RIGHT : DT_LEFT

        if !(dwStyle & BS_MULTILINE) ; || (dwStyle & BS_VCENTER)
            dwTextFlags |= DT_SINGLELINE
        
        center  := !!(dwTextFlags & DT_CENTER)
        vcenter := !!(dwTextFlags & DT_VCENTER)
        right   := !!(dwTextFlags & DT_RIGHT)
        bottom  := !!(dwTextFlags & DT_BOTTOM)
        
        return dwTextFlags | DT_WORDBREAK
    }

    /** @prop {Integer} TextColor Set/ Get the Button Text Color (RGB). (To set the text colour, you must have used `SetColor()`, `SetBackColor()`, or `BackColor` to set the background colour at least once beforehand.) */
    TextColor {
        Get => this.HasProp("_textColor") && _BtnColor.RgbToBgr(this._textColor)
        Set => this._textColor := _BtnColor.RgbToBgr(value)
    }

    /** @prop {Integer} BackColor Set/ Get the Button background Color (RGB). */
    BackColor {
        Get => this.HasProp("_clr") && _BtnColor.RgbToBgr(this._clr)
        Set {
            if !this.HasProp("_first")
                this.SetColor(value)
            else {
                b := _BtnColor
                this.opt("-Redraw")
                this._clr         := b.RgbToBgr(value)
                this._isDark      := b.IsColorDark(clr := b.RgbToBgr(this._clr))
                this._hoverColor  := b.RgbToBgr(b.BrightenColor(clr, this._isDark ? 5 : -5))
                this._pushedColor := b.RgbToBgr(b.BrightenColor(clr, this._isDark ? -10 : 10))
                this.opt("+Redraw")
            }
        }
    }

    /** @prop {Integer} BorderColor Button border color (RGB). (To set the border colour, you must have used `SetColor()`, `SetBackColor()`, or `BackColor` to set the background colour at least once beforehand.)*/
    BorderColor {
        Get => this.HasProp("_borderColor") && _BtnColor.RgbToBgr(this._borderColor)
        Set => this._borderColor := _BtnColor.RgbToBgr(value)
    }
    
    /** @prop {Integer} RoundedCorner Rounded corner preference for the button.  (To set the rounded corner preference, you must have used `SetColor()`, `SetBackColor()`, or `BackColor` to set the background colour at least once beforehand.) */
    RoundedCorner {
        Get => this.HasProp("_roundedCorner") && this._roundedCorner
        Set => this._roundedCorner := value
    }

    /**
     * @prop {Integer} ShowBorder
     * Border preference. (To set the border preference, you must have used `SetColor()`, `SetBackColor()`, or `BackColor` to set the background colour at least once beforehand.)
     * - `n` : The higher the value, the thicker the button's border when focused.
     * - `1` : Highlight when focused.
     * - `0` : No border displayed.
     * - `-1`: Border always visible.
     * - `-n`: The lower the value, the thicker the button's border when always visible.
     */
    ShowBorder {
        Get => this.HasProp("_showBorder") && this._showBorder
        Set {
            if IsNumber(Value)
                this._showBorder := value
            else throw TypeError("The value must be a number.", "ShowBorder")
        }
    }

    /**
     * Configures a button's appearance.
     * @param {number} bgColor - Button's background color (RGB).
     * @param {number} [colorBehindBtn] - Color of the button's surrounding area (defaults to `myGui.BackColor`).
     * @param {number} [roundedCorner] - Rounded corner preference for the button. If omitted, 
     * - For Windows 11: Enabled (value: 9).
     * - For Windows 10: Disabled.
     * @param {boolean} [showBorder=1]
     * - `n` : The higher the value, the thicker the button's border when focused.
     * - `1`: Highlight when focused.
     * - `0`: No border displayed.
     * - `-1`: Border always visible.
     * - `-n`: The lower the value, the thicker the button's border when always visible.
     * @param {number} [borderColor=0xFFFFFF] - Button border color (RGB).
     * @param {number} [txColor] - Button text color (RGB). If omitted, the text colour will be automatically set to white or black depends on the background colour.
     */
    SetBackColor(bgColor, colorBehindBtn?, roundedCorner?, showBorder := 1, borderColor := 0xFFFFFF, txColor?) => this.SetColor(bgColor, txColor?, showBorder, borderColor?, roundedCorner?)
    
    /**
     * Configures a button's appearance.
     * @param {number} bgColor - Button's background color (RGB).
     * @param {number} [txColor] - Button text color (RGB). If omitted, the text colour will be automatically set to white or black depends on the background colour.
     * @param {boolean} [showBorder=1]
     * - `n` : The higher the value, the thicker the button's border when focused.
     * - `1` : Highlight when focused.
     * - `0` : No border displayed.
     * - `-1`: Border always visible.
     * - `-n`: The lower the value, the thicker the button's border when always visible.
     * @param {number} [borderColor=0xFFFFFF] - Button border color (RGB).
     * @param {number} [roundedCorner] - Rounded corner preference for the button. If omitted,     
     * - For Windows 11: Enabled (value: 9).
     * - For Windows 10: Disabled.
     */
    SetColor(bgColor, txColor?, showBorder := 1, borderColor := 0xFFFFFF, roundedCorner?)
    { 
        static BS_BITMAP       := 0x0080
        static BS_FLAT         := 0x8000
        static IS_WIN11        := (VerCompare(A_OSVersion, "10.0.22200") >= 0)
        static NM_CUSTOMDRAW   := -12
        static WM_CTLCOLORBTN  := 0x0135
        static WS_CLIPSIBLINGS := 0x04000000
        static BTN_STYLE       := (WS_CLIPSIBLINGS | BS_FLAT | BS_BITMAP) 

        this._first         := 1
        this._roundedCorner := roundedCorner ?? (IS_WIN11 ? 9 : 0)
        this._showBorder    := showBorder
        this._clr           := ColorHex(bgColor)
        this._isDark        := _BtnColor.IsColorDark(this._clr)
        this._hoverColor    := _BtnColor.RgbToBgr(BrightenColor(this._clr, this._isDark ? 5 : -5))
        this._pushedColor   := _BtnColor.RgbToBgr(BrightenColor(this._clr, this._isDark ? -10 : 10))
        this._clr           := _BtnColor.RgbToBgr(this._clr)
        this._btnBkColor    := (colorBehindBtn ?? !IS_WIN11) && _BtnColor.RgbToBgr("0x" (this.Gui.BackColor))
        this._borderColor   := _BtnColor.RgbToBgr(borderColor)
        
        if !this.HasProp("_textColor") || IsSet(txColor)
            this._textColor := _BtnColor.RgbToBgr(txColor ?? (this._isDark ? 0xFFFFFF : 0))
        
        ; Uncomment the line blow if the button corner is a bit off.
        ; this.Gui.OnMessage(WM_CTLCOLORBTN, ON_WM_CTLCOLORBTN)

        if this._btnBkColor
            this.Gui.OnEvent("Close", (*) => (this.HasProp("__hbrush") ? DeleteObject(this.__hbrush) : 0))

        this.Opt(BTN_STYLE (IsSet(colorBehindBtn) ? " Background" colorBehindBtn : "")) ;  
        this.OnNotify(NM_CUSTOMDRAW, ON_NM_CUSTOMDRAW)

        if this._isDark
            SetWindowTheme(this.hwnd, "DarkMode_Explorer")

        SetWindowPos(this.hwnd, 0,,,,, 0x4043)
        this.Redraw()

        ON_NM_CUSTOMDRAW(gCtrl, lParam)
        {
            static CDDS_PREPAINT    := 0x1
            static CDIS_HOT         := 0x40
            static CDRF_DODEFAULT   := 0x0
            static CDRF_SKIPDEFAULT := 0x4
            static DC_BRUSH         := GetStockObject(18)
            static DC_PEN           := GetStockObject(19)
            static DT_CALCRECT      := 0x400
            static DT_WORDBREAK     := 0x10
            static PS_SOLID         := 0
            
            nmcd := NMCUSTOMDRAWINFO(lParam)

            if (nmcd.hdr.code != NM_CUSTOMDRAW 
            || nmcd.hdr.hwndFrom != gCtrl.hwnd
            || nmcd.dwDrawStage  != CDDS_PREPAINT)
                return CDRF_DODEFAULT
            
            ; Determine the background colour based on the button's status.
            isPressed := GetKeyState("LButton", "P")
            isHot     := (nmcd.uItemState & CDIS_HOT)
            brushColor := penColor := (!isHot || this._first ? this._clr : isPressed ? this._pushedColor : this._hoverColor)
            
            ; Set Rounded Corner Preference ----------------------------------------------

            rc     := nmcd.rc
            corner := this._roundedCorner
            SetWindowRgn(gCtrl.hwnd, CreateRoundRectRgn(rc.left, rc.top, rc.right, rc.bottom, corner, corner), 1)
            GetWindowRgn(gCtrl.hwnd, rcRgn := CreateRectRgn())
            
            ; Draw Border ----------------------------------------------------------------

            if ((this._showBorder < 0) || (this._showBorder > 0 && gCtrl.Focused)) {
                penColor := this._showBorder > 0 && !gCtrl.Focused ? penColor : this._borderColor
                hpen     := CreatePen(PS_SOLID, this._showBorder, penColor)
                SelectObject(nmcd.hdc, hpen)
                FrameRect(nmcd.hdc, rc, DC_PEN)                
            } else {
                SelectObject(nmcd.hdc, DC_PEN)
                SetDCPenColor(nmcd.hdc, penColor)
            }

            ; Draw Background ------------------------------------------------------------

            SelectObject(nmcd.hdc, DC_BRUSH)
            SetDCBrushColor(nmcd.hdc, brushColor)
            RoundRect(nmcd.hdc, rc.left, rc.top, rc.right-1, rc.bottom-1, corner, corner)

            ; Darw Text ------------------------------------------------------------------

            textPtr     := StrPtr(gCtrl.Text)
            dwTextFlags := this.GetTextFlags(&hCenter, &vCenter, &right, &bottom)
            SetBkMode(nmcd.hdc, 0)
            SetTextColor(nmcd.hdc, this._textColor)

            CopyRect(rcT := !NMCUSTOMDRAWINFO.HasProp("RECT") && IsSet(RECT) ? RECT() : NMCUSTOMDRAWINFO.RECT(), nmcd.rc)
            
            ; Calculate the text rect.
            DrawText(nmcd.hdc, textPtr, -1, rcT, DT_CALCRECT | dwTextFlags)

            if (hCenter || right)
                offsetW := ((nmcd.rc.width - rcT.Width - (right * 4)) / (hCenter ? 2 : 1))

            if (bottom || vCenter)
                offsetH := ((nmcd.rc.height - rct.Height - (bottom * 4)) / (vCenter ? 2 : 1))
                
            OffsetRect(rcT, offsetW ?? 2,offsetH ?? 2)
            DrawText(nmcd.hdc, textPtr, -1, rcT, dwTextFlags)

            if this._first
                this._first := 0

            DeleteObject(rcRgn)
            
            if (pen??0)
                DeleteObject(hpen)

            SetWindowPos(this.hwnd, 0, 0, 0, 0, 0, 0x4043)

            return CDRF_SKIPDEFAULT 
        }

        ON_WM_CTLCOLORBTN(GuiObj, wParam, lParam, Msg)
        {
            if (lParam != this.hwnd || !this.Focused)
                return

            SelectObject(wParam, hbrush := GetStockObject(18))
            SetBkMode(wParam, 0)

            if (colorBehindBtn ?? !IS_WIN11) {
                SetDCBrushColor(wParam, this._btnBkColor)
                SetBkColor(wParam, this._btnBkColor)
            }

            return hbrush 
        }

        BrightenColor(clr, perc := 5) => _BtnColor.BrightenColor(clr, perc)

        ColorHex(clr) => Number(((Type(clr) = "string" && SubStr(clr, 1, 2) != "0x") ? "0x" clr : clr))

        CopyRect(lprcDst, lprcSrc) => DllCall("CopyRect", "ptr", lprcDst, "ptr", lprcSrc, "int")

        CreateRectRgn(nLeftRect := 0, nTopRect := 0, nRightRect := 0, nBottomRect := 0) => DllCall('Gdi32\CreateRectRgn', 'int', nLeftRect, 'int', nTopRect, 'int', nRightRect, 'int', nBottomRect, 'ptr')

        CreateRoundRectRgn(nLeftRect, nTopRect, nRightRect, nBottomRect, nWidthEllipse, nHeightEllipse) => DllCall('Gdi32\CreateRoundRectRgn', 'int', nLeftRect, 'int', nTopRect, 'int', nRightRect, 'int', nBottomRect, 'int', nWidthEllipse, 'int', nHeightEllipse, 'ptr')

        CreatePen(fnPenStyle, nWidth, crColor) => DllCall('Gdi32\CreatePen', 'int', fnPenStyle, 'int', nWidth, 'uint', crColor, 'ptr')

        CreateSolidBrush(crColor) => DllCall('Gdi32\CreateSolidBrush', 'uint', crColor, 'ptr')

        DefWindowProc(hWnd, Msg, wParam, lParam) => DllCall("User32\DefWindowProc", "ptr", hWnd, "uint", Msg, "uptr", wParam, "uptr", lParam, "ptr")

        DeleteObject(hObject) => DllCall('Gdi32\DeleteObject', 'ptr', hObject, 'int')

        DrawText(hDC, lpchText, nCount, lpRect, uFormat) => DllCall("DrawText", "ptr", hDC, "ptr", lpchText, "int", nCount, "ptr", lpRect, "uint", uFormat, "int")

        FrameRect(hDC, lprc, hbr) => DllCall("FrameRect", "ptr", hDC, "ptr", lprc, "ptr", hbr, "int")

        FrameRgn(hdc, hrgn, hbr, nWidth, nHeight) => DllCall('Gdi32\FrameRgn', 'ptr', hdc, 'ptr', hrgn, 'ptr', hbr, 'int', nWidth, 'int', nHeight, 'int')

        GetStockObject(fnObject) => DllCall('Gdi32\GetStockObject', 'int', fnObject, 'ptr')

        GetWindowRgn(hWnd, hRgn, *) => DllCall("User32\GetWindowRgn", "ptr", hWnd, "ptr", hRgn, "int")

        OffsetRect(lprc, dx, dy) => DllCall("User32\OffsetRect", "ptr", lprc, "int", dx, "int", dy, "int")

        RGB(R := 255, G := 255, B := 255) => _BtnColor.RGB(R, G, B)

        RoundRect(hdc, nLeftRect, nTopRect, nRightRect, nBottomRect, nWidth, nHeight) => DllCall('Gdi32\RoundRect', 'ptr', hdc, 'int', nLeftRect, 'int', nTopRect, 'int', nRightRect, 'int', nBottomRect, 'int', nWidth, 'int', nHeight, 'int')

        SelectObject(hdc, hgdiobj) => DllCall('Gdi32\SelectObject', 'ptr', hdc, 'ptr', hgdiobj, 'ptr')

        SetBkColor(hdc, crColor) => DllCall('Gdi32\SetBkColor', 'ptr', hdc, 'uint', crColor, 'uint')

        SetBkMode(hdc, iBkMode) => DllCall('Gdi32\SetBkMode', 'ptr', hdc, 'int', iBkMode, 'int')

        SetDCBrushColor(hdc, crColor) => DllCall('Gdi32\SetDCBrushColor', 'ptr', hdc, 'uint', crColor, 'uint')

        SetDCPenColor(hdc, crColor) => DllCall('Gdi32\SetDCPenColor', 'ptr', hdc, 'uint', crColor, 'uint')

        SetTextColor(hdc, color) => DllCall("SetTextColor", "Ptr", hdc, "UInt", color)

        SetWindowPos(hWnd, hWndInsertAfter, X := 0, Y := 0, cx := 0, cy := 0, uFlags := 0x40) => DllCall("User32\SetWindowPos", "ptr", hWnd, "ptr", hWndInsertAfter, "int", X, "int", Y, "int", cx, "int", cy, "uint", uFlags, "int")

        SetWindowRgn(hWnd, hRgn, bRedraw) => DllCall("User32\SetWindowRgn", "ptr", hWnd, "ptr", hRgn, "int", bRedraw, "int")

        SetWindowTheme(hwnd, appName, subIdList?) => DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "ptr", StrPtr(appName), "ptr", subIdList ?? 0)
    }

    static RGB(R := 255, G := 255, B := 255) => ((R << 16) | (G << 8) | B)

    static BrightenColor(clr, perc := 5) => ((p := perc / 100 + 1), _BtnColor.RGB(Round(Min(255, (clr >> 16 & 0xFF) * p)), Round(Min(255, (clr >> 8 & 0xFF) * p)), Round(Min(255, (clr & 0xFF) * p))))
    
    static IsColorDark(clr) => (((clr >> 16 & 0xFF) / 255 * 0.2126 + (clr >> 8 & 0xFF) / 255 * 0.7152 + (clr & 0xFF) / 255 * 0.0722) < 0.5)

    static RgbToBgr(color) => (Type(color) = "string") ? this.RgbToBgr(Number(SubStr(Color, 1, 2) = "0x" ? color : "0x" color)) : (Color >> 16 & 0xFF) | (Color & 0xFF00) | ((Color & 0xFF) << 16)
}

; ========================================= DARK MSGBOX =========================================


#DllLoad gdi32.dll

class DarkMsgBox
{
    static __New()
    {
        /** Thanks to geekdude & Mr Doge for providing this method to rewrite built-in functions. */
        static _Msgbox   := MsgBox.Call.Bind(MsgBox)
        static _InputBox := InputBox.Call.Bind(InputBox)
        MsgBox.DefineProp("Call", {Call: CallNativeFunc})
        InputBox.DefineProp("Call", {Call: CallNativeFunc})

        CallNativeFunc(_this, params*)
        {
            static WM_COMMNOTIFY := 0x44
            static WM_INITDIALOG := 0x0110
            
            iconNumber := 1
            iconFile   := ""
            
            if (params.length = (_this.MaxParams + 2))
                iconNumber := params.Pop()
            
            if (params.length = (_this.MaxParams + 1)) 
                iconFile := params.Pop()
            
            SetThreadDpiAwarenessContext(-3)
    
            if InStr(_this.Name, "MsgBox")
                OnMessage(WM_COMMNOTIFY, ON_WM_COMMNOTIFY)
            else
                OnMessage(WM_INITDIALOG, ON_WM_INITDIALOG, -1)
    
            return _%_this.name%(params*)
    
            ON_WM_INITDIALOG(wParam, lParam, msg, hwnd)
            {
                OnMessage(WM_INITDIALOG, ON_WM_INITDIALOG, 0)
                WNDENUMPROC(hwnd)
            }
            
            ON_WM_COMMNOTIFY(wParam, lParam, msg, hwnd)
            {
                if (msg = 68 && wParam = 1027)
                    OnMessage(0x44, ON_WM_COMMNOTIFY, 0),                    
                    EnumThreadWindows(GetCurrentThreadId(), CallbackCreate(WNDENUMPROC), 0)
            }
    
            WNDENUMPROC(hwnd, *)
            {
                global currentTheme
                
                ; Only apply dark styling if the current theme is dark
                if (currentTheme != "dark")
                    return 0
                    
                static SM_CICON         := "W" SysGet(11) " H" SysGet(12)
                static SM_CSMICON       := "W" SysGet(49) " H" SysGet(50)
                static ICON_BIG         := 1
                static ICON_SMALL       := 0
                static WM_SETICON       := 0x80
                static WS_CLIPCHILDREN  := 0x02000000
                static WS_CLIPSIBLINGS  := 0x04000000
                static WS_EX_COMPOSITED := 0x02000000
                static winAttrMap       := Map(10, true, 17, true, 20, true, 38, 4, 35, 0x2b2b2b)
    
                SetWinDelay(-1)
                SetControlDelay(-1)
                DetectHiddenWindows(true)
    
                if !WinExist("ahk_class #32770 ahk_id" hwnd)
                    return 1
    
                WinSetStyle("+" (WS_CLIPSIBLINGS | WS_CLIPCHILDREN))
                WinSetExStyle("+" (WS_EX_COMPOSITED))
                SetWindowTheme(hwnd, "DarkMode_Explorer")
    
                if iconFile {
                    hICON_SMALL := LoadPicture(iconFile, SM_CSMICON " Icon" iconNumber, &handleType)
                    hICON_BIG   := LoadPicture(iconFile, SM_CICON " Icon" iconNumber, &handleType)
                    PostMessage(WM_SETICON, ICON_SMALL, hICON_SMALL)
                    PostMessage(WM_SETICON, ICON_BIG, hICON_BIG)
                }
    
                for dwAttribute, pvAttribute in winAttrMap
                    DwmSetWindowAttribute(hwnd, dwAttribute, pvAttribute)
                
                GWL_WNDPROC(hwnd, hICON_SMALL?, hICON_BIG?)
                return 0
            }
            
            GWL_WNDPROC(winId := "", hIcons*)
            {
                global currentTheme
                
                ; Only proceed with dark styling if the current theme is dark
                if (currentTheme != "dark")
                    return
                    
                static SetWindowLong     := DllCall.Bind(A_PtrSize = 8 ? "SetWindowLongPtr" : "SetWindowLong", "ptr",, "int",, "ptr",, "ptr")
                static BS_FLAT           := 0x8000
                static BS_BITMAP         := 0x0080
                static DPI               := (A_ScreenDPI / 96)
                static WM_CLOSE          := 0x0010
                static WM_CTLCOLORBTN    := 0x0135
                static WM_CTLCOLORDLG    := 0x0136
                static WM_CTLCOLOREDIT   := 0x0133
                static WM_CTLCOLORSTATIC := 0x0138
                static WM_DESTROY        := 0x0002
                static WM_SETREDRAW      := 0x000B
    
                SetControlDelay(-1)
    
                btns    := []
                btnHwnd := ""
    
                for ctrl in WinGetControlsHwnd(winId)
                {
                    classNN := ControlGetClassNN(ctrl)
                    SetWindowTheme(ctrl, !InStr(classNN, "Edit") ? "DarkMode_Explorer" : "DarkMode_CFD")
    
                    if InStr(classNN, "B") 
                        btns.Push(btnHwnd := ctrl)
                }
    
                WindowProcOld := SetWindowLong(winId, -4, CallbackCreate(WNDPROC))
                
                WNDPROC(hwnd, uMsg, wParam, lParam)
                {
                    static hbrush := []
                    SetWinDelay(-1)
                    SetControlDelay(-1)
                    
                    if !hbrush.Length
                        for clr in [0x202020, 0x2b2b2b]
                            hbrush.Push(CreateSolidBrush(clr))
    
                    switch uMsg {
                    case WM_CTLCOLORSTATIC: 
                    {
                        SelectObject(wParam, hbrush[2])
                        SetBkMode(wParam, 0)
                        SetTextColor(wParam, 0xFFFFFF)
                        SetBkColor(wParam, 0x2b2b2b)
    
                        for _hwnd in btns
                            PostMessage(WM_SETREDRAW,,,_hwnd)
    
                        GetClientRect(winId, rcC := this.RECT())
                        WinGetClientPos(&winX, &winY, &winW, &winH, winId)
                        ControlGetPos(, &btnY,, &btnH, btnHwnd)
                        hdc        := GetDC(winId)
                        rcC.top    := btnY - (rcC.bottom - (btnY+btnH))
                        rcC.bottom *= 2
                        rcC.right  *= 2
                        
                        SetBkMode(hdc, 0)
                        FillRect(hdc, rcC, hbrush[1])
                        ReleaseDC(winId, hdc)
    
                        for _hwnd in btns
                            PostMessage(WM_SETREDRAW, 1,,_hwnd)
    
                        return hbrush[2]
                    }
                    case WM_CTLCOLORBTN, WM_CTLCOLORDLG, WM_CTLCOLOREDIT: 
                    {         
                        brushIndex := !(uMsg = WM_CTLCOLORBTN)
                        SelectObject(wParam, brush := hbrush[brushIndex+1])
                        SetBkMode(wParam, 0)
                        SetTextColor(wParam, 0xFFFFFF)
                        SetBkColor(wParam, !brushIndex ? 0x202020 : 0x2b2b2b)
                        return brush
                    }
                    case WM_DESTROY: 
                    {
                        for v in hIcons
                            (v??0) && DestroyIcon(v)
    
                        while hbrush.Length
                            DeleteObject(hbrush.Pop())
                    }}
    
                    return CallWindowProc(WindowProcOld, hwnd, uMsg, wParam, lParam) 
                }
            }
    
            CreateSolidBrush(crColor) => DllCall('Gdi32\CreateSolidBrush', 'uint', crColor, 'ptr')
            
            CallWindowProc(lpPrevWndFunc, hWnd, uMsg, wParam, lParam) => DllCall("CallWindowProc", "Ptr", lpPrevWndFunc, "Ptr", hwnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam)
    
            DestroyIcon(hIcon) => DllCall("DestroyIcon", "ptr", hIcon)
    
            /** @see — https://learn.microsoft.com/en-us/windows/win32/api/dwmapi/ne-dwmapi-dwmwindowattribute */
            DWMSetWindowAttribute(hwnd, dwAttribute, pvAttribute, cbAttribute := 4) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", dwAttribute, "Ptr*", &pvAttribute, "UInt", cbAttribute)
            
            DeleteObject(hObject) => DllCall('Gdi32\DeleteObject', 'ptr', hObject, 'int')
            
            EnumThreadWindows(dwThreadId, lpfn, lParam) => DllCall("User32\EnumThreadWindows", "uint", dwThreadId, "ptr", lpfn, "uptr", lParam, "int")
            
            FillRect(hDC, lprc, hbr) => DllCall("User32\FillRect", "ptr", hDC, "ptr", lprc, "ptr", hbr, "int")
            
            GetClientRect(hWnd, lpRect) => DllCall("User32\GetClientRect", "ptr", hWnd, "ptr", lpRect, "int")
            
            GetCurrentThreadId() => DllCall("Kernel32\GetCurrentThreadId", "uint")
            
            GetDC(hwnd := 0) => DllCall("GetDC", "ptr", hwnd, "ptr")
    
            ReleaseDC(hWnd, hDC) => DllCall("User32\ReleaseDC", "ptr", hWnd, "ptr", hDC, "int")
            
            SelectObject(hdc, hgdiobj) => DllCall('Gdi32\SelectObject', 'ptr', hdc, 'ptr', hgdiobj, 'ptr')
            
            SetBkColor(hdc, crColor) => DllCall('Gdi32\SetBkColor', 'ptr', hdc, 'uint', crColor, 'uint')
            
            SetBkMode(hdc, iBkMode) => DllCall('Gdi32\SetBkMode', 'ptr', hdc, 'int', iBkMode, 'int')
    
            SetTextColor(hdc, crColor) => DllCall('Gdi32\SetTextColor', 'ptr', hdc, 'uint', crColor, 'uint')
            
            SetThreadDpiAwarenessContext(dpiContext) => DllCall("SetThreadDpiAwarenessContext", "ptr", dpiContext, "ptr")
    
            SetWindowTheme(hwnd, pszSubAppName, pszSubIdList := "") => (!DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "ptr", StrPtr(pszSubAppName), "ptr", pszSubIdList ? StrPtr(pszSubIdList) : 0) ? true : false)
        }
    }

    class RECT extends Buffer {
        static ofst := Map("left", 0, "top", 4, "right", 8, "bottom", 12)

        __New(left := 0, top := 0, right := 0, bottom := 0) {
            super.__New(16)
            NumPut("int", left, "int", top, "int", right, "int", bottom, this)
        }

        __Set(Key, Params, Value) {
            if DarkMsgBox.RECT.ofst.Has(k := StrLower(key))
                NumPut("int", value, this, DarkMsgBox.RECT.ofst[k])
            else throw PropertyError
        }

        __Get(Key, Params) {
            if DarkMsgBox.RECT.ofst.Has(k := StrLower(key))
                return NumGet(this, DarkMsgBox.RECT.ofst[k], "int")
            throw PropertyError
        }

        width  => this.right - this.left
        height => this.bottom - this.top
    }
} 

; ========================================= TOOLTIP =========================================

class SystemThemeAwareToolTip
{
    static GetIsDarkMode() {
        global currentTheme
        return currentTheme = "dark"
    }

    static IsDarkMode => SystemThemeAwareToolTip.GetIsDarkMode()

    static __New()
    {
        if this.HasOwnProp("HTT") || !this.IsDarkMode
            return

        GroupAdd("tooltips_class32", "ahk_class tooltips_class32")

        this.HTT        := DllCall("User32.dll\CreateWindowEx", "UInt", 8, "Ptr", StrPtr("tooltips_class32"), "Ptr", 0, "UInt", 3, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", A_ScriptHwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0)
        this.SubWndProc := CallbackCreate(TT_WNDPROC,, 4)
        this.OriWndProc := DllCall(A_PtrSize = 8 ? "SetClassLongPtr" : "SetClassLongW", "Ptr", this.HTT, "Int", -24, "Ptr", this.SubWndProc, "UPtr")
        
        TT_WNDPROC(hWnd, uMsg, wParam, lParam)
        {
            static WM_CREATE := 0x0001
            global currentTheme
            
            if (currentTheme = "dark" && uMsg = WM_CREATE)
            {
                SetDarkToolTip(hWnd)

                if (VerCompare(A_OSVersion, "10.0.22000") > 0)
                    SetRoundedCornor(hWnd, 3)
            }

            return DllCall(This.OriWndProc, "Ptr", hWnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam, "UInt")
        }

        SetDarkToolTip(hWnd) => DllCall("UxTheme\SetWindowTheme", "Ptr", hWnd, "Ptr", StrPtr("DarkMode_Explorer"), "Ptr", StrPtr("ToolTip"))

        SetRoundedCornor(hwnd, level:= 3) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", 33, "Ptr*", level, "UInt", 4)
    }

    static __Delete() => (this.HTT && WinKill("ahk_group tooltips_class32"))
}