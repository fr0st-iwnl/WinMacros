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

global currentVersion := "1.3"
global versionCheckUrl := "https://winmacros.netlify.app/version/version.txt"
global githubReleasesUrl := "https://github.com/fr0st-iwnl/WinMacros/releases/latest"

global currentTheme := IniRead(settingsFile, "Settings", "Theme", "dark")
global notificationsEnabled := IniRead(settingsFile, "Settings", "Notifications", "1") = "1"

global unifiedGui := ""
global currentTabIndex := 1

global currentSetHotkeyGui := ""
global currentSetHotkeyAction := ""

global powerGui := ""

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
    A_TrayMenu.Delete()
    
    A_TrayMenu.Add("Open WinMacros", (*) => ShowUnifiedGUI(1))
    A_TrayMenu.Add()
    A_TrayMenu.Add("Show Welcome Message at startup", ToggleStartup)
    A_TrayMenu.Add("Run on Windows Startup", ToggleWindowsStartup)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Enable Notifications", ToggleNotifications)
    A_TrayMenu.Add("Dark Theme", ToggleTrayTheme)
    A_TrayMenu.Add("Check for Updates", (*) => CheckForUpdates(true))
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
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
}

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
    A_TrayMenu.Add("Exit", (*) => ExitApp())
}

InitializeTrayMenu()

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
    global notificationQueue, isShowingNotification, currentTheme, activeNotifications
    
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
    
    activeNotifications.Push({gui: notify, height: height, text: message, isSpecial: isSpecial})
    
    SetTimer(RemoveNotification.Bind(notify, xPos, top), -2000)
    
    notificationQueue.RemoveAt(1)
    isShowingNotification := false
    
    if (notificationQueue.Length > 0) {
        SetTimer(ShowNextNotification, -10)
    }
}

RemoveNotification(notify, xPos, topPos) {
    global activeNotifications
    
    for i, notifyInfo in activeNotifications {
        if (notifyInfo.gui = notify) {
            activeNotifications.RemoveAt(i)
            break
        }
    }
    
    notify.Destroy()
    
    yPos := topPos + 20
    for i, notifyInfo in activeNotifications {
        notifyInfo.gui.Move(xPos, yPos)
        yPos += notifyInfo.height + 10
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
    global currentSetHotkeyGui, currentSetHotkeyAction, unifiedGui
    static setHotkeyGuiOpen := false
    
    if (setHotkeyGuiOpen) {
        return
    }
    
    setHotkeyGuiOpen := true
    
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
            inputGui.Destroy()
            setHotkeyGuiOpen := false
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
            inputGui.Destroy()
            setHotkeyGuiOpen := false
            if (IsObject(unifiedGui))
                unifiedGui.Destroy()
            ShowUnifiedGUI(2)
        }
    }
    
    okBtn := currentSetHotkeyGui.Add("Button", "x10 w100", "OK")
    if (currentTheme = "dark") {
        okBtn.SetFont("c0xFFFFFF")
        okBtn.Opt("+Background333333")
    } else {
        okBtn.SetFont("c0x000000")
        okBtn.Opt("+BackgroundDDDDDD")
    }
    okBtn.OnEvent("Click", (*) => (SetHotkey(currentSetHotkeyGui, action), setHotkeyGuiOpen := false))
    
    cancelBtn := currentSetHotkeyGui.Add("Button", "x+10 w100", "Cancel")
    if (currentTheme = "dark") {
        cancelBtn.SetFont("c0xFFFFFF")
        cancelBtn.Opt("+Background333333")
    } else {
        cancelBtn.SetFont("c0x000000")
        cancelBtn.Opt("+BackgroundDDDDDD")
    }
    cancelBtn.OnEvent("Click", (*) => (currentSetHotkeyGui.Destroy(), setHotkeyGuiOpen := false))
    
    currentSetHotkeyGui.OnEvent("Close", (*) => (currentSetHotkeyGui.Destroy(), setHotkeyGuiOpen := false))
    
    currentSetHotkeyGui.Show()
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
    global isMenuOpen, currentTheme, powerGui
    
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
    global isEditorMenuOpen, editorGui, currentSelection
    
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
    global isEditorMenuOpen, editorGui
    
    try {
        Hotkey "Up", "Off"
        Hotkey "Down", "Off"
        Hotkey "Enter", "Off"
        Hotkey "Escape", "Off"
    } catch Error {
    }
    
    try {
        if IsObject(editorGui) {
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

ToggleTheme(isDark) {
    global currentTheme, currentSetHotkeyGui, currentSetHotkeyAction, isMenuOpen, powerGui, unifiedGui
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
    
    if (IsObject(unifiedGui)) {
        currentTab := unifiedGui["TabControl"].Value
        unifiedGui.Destroy()
        ShowUnifiedGUI(currentTab)
    }
    
    if (IsObject(currentSetHotkeyGui) && WinExist("Set Hotkey")) {
        try {
            currentHotkey := currentSetHotkeyGui["NewHotkey"].Value
            currentSetHotkeyGui.Destroy()
            SetNewHotkeyGUI(currentSetHotkeyAction, "")
            currentSetHotkeyGui["NewHotkey"].Value := currentHotkey
        }
    }
    
    if (IsObject(powerGui)) {
        powerGui.Destroy()
        isMenuOpen := false
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
    
    tabs := unifiedGui.Add("Tab3", "vTabControl w770 h600", ["👋 Welcome", "⌨️ Hotkey Settings", "🚀 Launcher Settings", "🛠️ General Settings"])
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
    
    unifiedGui.Show("w800 h650 Center")
    
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
    gui.Add("Text", "x20 y40 w200", "Action")
    gui.Add("Text", "x310 y40 w200", "Current Hotkey")
    gui.Add("Text", "x545 y40 w100", "Modify")
    
    y := 70
    
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "Applications")
    y += 30
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("OpenExplorer", y, gui)
    y += 30
    AddHotkeyRowToTab("OpenPowerShell", y, gui)
    y += 30
    AddHotkeyRowToTab("OpenBrowser", y, gui)
    y += 30
    AddHotkeyRowToTab("OpenVSCode", y, gui)
    y += 30
    AddHotkeyRowToTab("OpenCalculator", y, gui)
    y += 30
    AddHotkeyRowToTab("OpenSpotify", y, gui)
    
    y += 40
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "System Tools")
    y += 30
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("ToggleTaskbar", y, gui)
    y += 30
    AddHotkeyRowToTab("ToggleDesktopIcons", y, gui)
    
    y += 40
    gui.Add("Text", "x20 y" y " w730 h2 0x10")
    y += 10
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    gui.Add("Text", "x20 y" y " w200", "Sound Controls")
    y += 30
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "White" : "Black"), "Segoe UI")
    
    AddHotkeyRowToTab("VolumeUp", y, gui)
    y += 30
    AddHotkeyRowToTab("VolumeDown", y, gui)
    y += 30
    AddHotkeyRowToTab("ToggleMute", y, gui)
    y += 30
    AddHotkeyRowToTab("ToggleMic", y, gui)
    
    resetBtn := gui.Add("Button", "x650 y535 w100 h30", "Reset All")
    resetBtn.SetFont("c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"))
    resetBtn.Opt("+Background" (currentTheme = "dark" ? "333333" : "DDDDDD"))
    resetBtn.OnEvent("Click", (*) => ResetAllHotkeysTab())
}

AddHotkeyRowToTab(action, y, gui) {
    global currentTheme, hotkeyActions, keybindsFile
    currentHotkey := IniRead(keybindsFile, "Hotkeys", action, "None")
    
    gui.Add("Text", "x" (gui.MarginX + 20) " y" y " w130 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), hotkeyActions[action])
    
    textColor := "c0x000000"
    bgColor := currentHotkey = "None" ? "D3D3D3" : "98FB98"
    hotkeyText := gui.Add("Text", "x+95 y" y " w200 h25 Center Border " textColor " Background" bgColor,
        FormatHotkey(currentHotkey))
    
    CreateButtonInTab(gui, action, y)
}

CreateButtonInTab(gui, action, y) {
    global currentTheme
    
    btn := gui.Add("Button", "x+60 y" y " w100 h25 -E0x200", "Set Hotkey")
    if (currentTheme = "dark") {
        btn.SetFont("c0xFFFFFF")
        btn.Opt("+Background333333")
    } else {
        btn.SetFont("c0x000000")
        btn.Opt("+BackgroundDDDDDD")
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
    
    gui.Add("Text", "x20 y380 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Name:")
    nameInput := gui.Add("Edit", "x120 y380 w200 h23 vLauncherName " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    
    gui.Add("Text", "x20 y410 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Hotkey:")
    hotkeyInput := gui.Add("Hotkey", "x120 y410 w200 vLauncherHotkey")
    helpText := gui.Add("Link", "x+10 y413 w20 h20 -TabStop", '<a href="#">?</a>')
    helpText.SetFont("bold s10 underline c" (currentTheme = "dark" ? "0x98FB98" : "Blue"), "Segoe UI")
    helpText.OnEvent("Click", ShowLauncherHotkeyTooltip)
    
    gui.OnEvent("Escape", (*) => hotkeyInput.Value := "")

    gui.Add("Text", "x20 y440 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Program Path:")
    pathInput := gui.Add("Edit", "x120 y440 w500 h23 vLauncherPath " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    browseBtn := gui.Add("Button", "x630 y438 w100", "Browse")
    
    addBtn := gui.Add("Button", "x120 y480 w100", "Add")
    editBtn := gui.Add("Button", "x230 y480 w100", "Edit")
    deleteBtn := gui.Add("Button", "x340 y480 w100", "Delete")
    
    if (currentTheme = "dark") {
        browseBtn.Opt("+Background333333 cWhite")
        addBtn.Opt("+Background333333 cWhite")
        editBtn.Opt("+Background333333 cWhite")
        deleteBtn.Opt("+Background333333 cWhite")
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
    global unifiedGui, launcherIniPath, currentTheme
    static editGuiOpen := false
    
    if (editGuiOpen)
        return
        
    try {
        lv := unifiedGui["SysListView321"]
        if !(rowNum := lv.GetNext()) {
            ShowNotification("❌ Please select a launcher to edit")
            return
        }
        
        editGuiOpen := true
        
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
        
        editGui := Gui("+Owner" unifiedGui.Hwnd, "Edit Launcher")
        editGui.SetFont("s10", "Segoe UI")
        editGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "FFFFFF"
        
        if (currentTheme = "dark") {
            DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", editGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
            DllCall("uxtheme\SetWindowTheme", "Ptr", editGui.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
        }
        
        editGui.Add("Text", "x10 y10 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Name:")
        editNameInput := editGui.Add("Edit", "x110 y10 w200 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), oldName)
        
        editGui.Add("Text", "x10 y40 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Hotkey:")
        editHotkeyInput := editGui.Add("Hotkey", "x110 y40 w200", oldHotkey)
        
        editGui.OnEvent("Escape", (*) => editHotkeyInput.Value := "")

        editGui.Add("Text", "x10 y70 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Program Path:")
        editPathInput := editGui.Add("Edit", "x110 y70 w400 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), oldPath)
        editBrowseBtn := editGui.Add("Button", "x520 y68 w100", "Browse")
        
        saveBtn := editGui.Add("Button", "x420 y110 w100", "Save")
        cancelBtn := editGui.Add("Button", "x530 y110 w100", "Cancel")
        
        if (currentTheme = "dark") {
            editBrowseBtn.Opt("+Background333333 cWhite")
            saveBtn.Opt("+Background333333 cWhite")
            cancelBtn.Opt("+Background333333 cWhite")
        }
        
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
            
            editGui.Destroy()
            editGuiOpen := false
            
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
        cancelBtn.OnEvent("Click", (*) => (editGui.Destroy(), editGuiOpen := false))
        editGui.OnEvent("Close", (*) => (editGui.Destroy(), editGuiOpen := false))
        
        editGui.Show("w640 h150")
    } catch Error as err {
        editGuiOpen := false
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
    global currentTheme, settingsFile, notificationsEnabled, currentVersion
    
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
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y160 w200", "Startup Options")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    showAtStartup := IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1"
    welcomeChk := gui.Add("Checkbox", "x" (leftColX+20) " y190 vShowWelcome", "Show welcome screen at startup")
    welcomeChk.Value := showAtStartup
    welcomeChk.OnEvent("Click", ToggleWelcomeStartup)
    
    isStartupEnabled := FileExist(A_Startup "\WinMacros.lnk") ? true : false
    startupChk := gui.Add("Checkbox", "x" (leftColX+20) " y220 vRunOnStartup", "Run WinMacros on Windows startup")
    startupChk.Value := isStartupEnabled
    startupChk.OnEvent("Click", ToggleWindowsStartupFromSettings)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y260 w200", "Notifications")
    
    gui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    notifyChk := gui.Add("Checkbox", "x" (leftColX+20) " y290 vEnableNotifications", "Enable notifications")
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
        exportHotkeysBtn.Opt("+Background333333 cWhite")
        importHotkeysBtn.Opt("+Background333333 cWhite")
        exportLaunchersBtn.Opt("+Background333333 cWhite")
        importLaunchersBtn.Opt("+Background333333 cWhite")
        openConfigBtn.Opt("+Background333333 cWhite")
    } else {
        exportHotkeysBtn.Opt("+BackgroundDDDDDD")
        importHotkeysBtn.Opt("+BackgroundDDDDDD")
        exportLaunchersBtn.Opt("+BackgroundDDDDDD")
        importLaunchersBtn.Opt("+BackgroundDDDDDD")
        openConfigBtn.Opt("+BackgroundDDDDDD")
    }
    
    exportHotkeysBtn.OnEvent("Click", ExportHotkeys)
    importHotkeysBtn.OnEvent("Click", ImportHotkeys)
    exportLaunchersBtn.OnEvent("Click", ExportLaunchers)
    importLaunchersBtn.OnEvent("Click", ImportLaunchers)
    openConfigBtn.OnEvent("Click", OpenConfigLocation)
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" rightColX " y230 w200", "Updates")
    
    checkUpdateBtn := gui.Add("Button", "x" (rightColX+20) " y260 w200 h30", "Check for Updates")
    checkUpdateBtn.SetFont("c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"))
    checkUpdateBtn.Opt("+Background" (currentTheme = "dark" ? "333333" : "DDDDDD"))
    checkUpdateBtn.OnEvent("Click", (*) => CheckForUpdates(true))
    
    gui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" leftColX " y330 w200", "About")
    
    gui.SetFont("s10 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    gui.Add("Text", "x" (leftColX+20) " y360", "[#] Current Version: " currentVersion)
    
    gui.SetFont("s10", "Segoe UI")
    linkColor := currentTheme = "dark" ? "cWhite" : "cBlue"
    githubLink := gui.Add("Link", "x" (leftColX+20) " y390 w300 -TabStop " linkColor, 
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