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

global currentVersion := "1.0"
global versionCheckUrl := "https://winmacros.netlify.app/version/version.txt"
global githubReleasesUrl := "https://github.com/fr0st-iwnl/WinMacros/releases"

global currentTheme := IniRead(settingsFile, "Settings", "Theme", "light")

global welcomeGui := ""

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
    IniWrite("light", settingsFile, "Settings", "Theme")
}

global launcherIniPath := EnvGet("LOCALAPPDATA") "\WinMacros\launcher.ini"
global activeHotkeys := Map()

InitializeTrayMenu() {
    A_TrayMenu.Delete()
    
    A_TrayMenu.Add("Show Welcome Screen", ShowWelcomeGUI)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Show Welcome Message at startup", ToggleStartup)
    A_TrayMenu.Add("Run on Windows Startup", ToggleWindowsStartup)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Open Hotkey Settings", (*) => ShowKeybindsGUI())
    A_TrayMenu.Add("Open Launcher Settings", (*) => ShowLauncherGUI())
    A_TrayMenu.Add()
    A_TrayMenu.Add("Dark Theme", ToggleTrayTheme)
    A_TrayMenu.Add("Check for Updates", (*) => CheckForUpdates())
    A_TrayMenu.Add("Exit", (*) => ExitApp())
    
    if (currentTheme = "dark")
        A_TrayMenu.Check("Dark Theme")
    if (IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1")
        A_TrayMenu.Check("Show Welcome Message at startup")
    if (FileExist(A_Startup "\WinMacros.lnk"))
        A_TrayMenu.Check("Run on Windows Startup")
}

ToggleStartup(*) {
    global settingsFile, welcomeGui
    isChecked := IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1"
    if (!isChecked) {
        IniWrite("1", settingsFile, "Settings", "ShowWelcome")
        ShowNotification("✅ Welcome screen enabled at startup")
        A_TrayMenu.Check("Show Welcome Message at startup")
        
        if (IsObject(welcomeGui) && WinExist("WinMacros: Welcome")) {
            try welcomeGui["ShowAtStartup"].Value := true
        }
    } else {
        IniWrite("0", settingsFile, "Settings", "ShowWelcome")
        ShowNotification("❌ Welcome screen disabled at startup")
        A_TrayMenu.Uncheck("Show Welcome Message at startup")
        
        if (IsObject(welcomeGui) && WinExist("WinMacros: Welcome")) {
            try welcomeGui["ShowAtStartup"].Value := false
        }
    }
}

ToggleWindowsStartup(*) {
    global welcomeGui
    startupPath := A_Startup "\WinMacros.lnk"
    
    if (!FileExist(startupPath)) {
        try {
            FileCreateShortcut(A_ScriptFullPath, startupPath,, "Launch WinMacros on startup")
            ShowNotification("✅ WinMacros will run on Windows startup")
            A_TrayMenu.Check("Run on Windows Startup")
            
            if (IsObject(welcomeGui) && WinExist("WinMacros: Welcome")) {
                try welcomeGui["RunOnStartup"].Value := true
            }
        } catch Error as err {
            ShowNotification("❌ Failed to create startup shortcut")
        }
    } else {
        try {
            FileDelete(startupPath)
            ShowNotification("❌ WinMacros will not run on Windows startup")
            A_TrayMenu.Uncheck("Run on Windows Startup")
            
            if (IsObject(welcomeGui) && WinExist("WinMacros: Welcome")) {
                try welcomeGui["RunOnStartup"].Value := false
            }
        } catch Error as err {
            ShowNotification("❌ Failed to remove startup shortcut")
        }
    }
}

ToggleTrayTheme(*) {
    global currentTheme
    ToggleTheme(currentTheme = "light")
}

CreateTrayMenu() {
    A_TrayMenu.Delete()
    A_TrayMenu.Add("Show Welcome Screen", ShowWelcomeGUI)
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
    SoundSetMute(-1)
    isMuted := SoundGetMute()
    ShowNotification(isMuted ? "🔇 Volume Muted" : "🔊 Volume Unmuted")
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

ToggleMic(*) {
    SoundSetMute(-1, , "Microphone")
    isMuted := SoundGetMute(, "Microphone")
    ShowNotification(isMuted ? "🎤 Microphone Muted" : "🎤 Microphone Unmuted")
}

OpenBrowser(*) {
    Run("http://")
    ShowNotification("🌐 Opening Default Browser")
}

VolumeUp(*) {
    SoundSetVolume("+5")
    currentVol := SoundGetVolume()
    ShowNotification("🔊 Volume: " Round(currentVol) "%")
}

VolumeDown(*) {
    SoundSetVolume("-5")
    currentVol := SoundGetVolume()
    ShowNotification("🔉 Volume: " Round(currentVol) "%")
}

ShowNotification(message) {
    global notificationQueue, isShowingNotification
    notificationQueue.Push(message)
    
    if (!isShowingNotification) {
        ShowNextNotification()
    }
}

ShowNextNotification() {
    global notificationQueue, isShowingNotification, currentNotify, currentTheme
    
    if (notificationQueue.Length = 0) {
        isShowingNotification := false
        return
    }
    
    message := notificationQueue[1]
    
    if (IsObject(currentNotify)) {
        try currentNotify.Destroy()
    }
    
    currentNotify := Gui("-Caption +AlwaysOnTop +ToolWindow")
    currentNotify.SetFont("s10", "Segoe UI")
    currentNotify.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"
    
    MonitorGetWorkArea(, &left, &top, &right, &bottom)
    
    if (StrLen(message) > 40) {
        height := 50
        currentNotify.Add("Text", "x10 y10 w280 r2 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), message)
    } else {
        currentNotify.Add("Text", "x10 y10 w280 r1 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), message)
        height := 35
    }
    
    width := 300
    xPos := right - width - 20
    yPos := top + 20
    
    currentNotify.Show(Format("NoActivate x{1} y{2} w{3} h{4}", xPos, yPos, width, height))
    
    isShowingNotification := true
    
    SetTimer(ProcessNextNotification, -2000)
}

ProcessNextNotification() {
    global currentNotify, notificationQueue, isShowingNotification
    
    if (IsObject(currentNotify)) {
        try currentNotify.Destroy()
    }
    currentNotify := ""
    
    if (notificationQueue.Length > 0) {
        notificationQueue.RemoveAt(1)
    }
    
    isShowingNotification := false
    
    if (notificationQueue.Length > 0) {
        SetTimer(ShowNextNotification, -10)
    }
}

^!k::ShowKeybindsGUI()

ShowKeybindsGUI(*) {
    global isMenuOpen
    
    if (isMenuOpen)
        return
        
    isMenuOpen := true
    
    static currentGui := ""
    
    if (IsObject(currentGui)) {
        currentGui.Destroy()
    }
    
    currentGui := Gui("+MinSize400x400", "Hotkey Settings")
    
    HotIfWinActive("Hotkey Settings")
    Hotkey "Escape", (*) => (isMenuOpen := false, currentGui.Destroy(), Hotkey("Escape", "Off")), "On"
    HotIfWinActive()
    
    currentGui.SetFont("s10", "Segoe UI")
    currentGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"

    if (currentTheme = "dark") {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", currentGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
    }
    
    currentGui.Add("Text", "x20 y20 w200 section c" (currentTheme = "dark" ? "White" : "Black"), "Action").SetFont("bold")
    currentGui.Add("Text", "x+95 w150 c" (currentTheme = "dark" ? "White" : "Black"), "Current Hotkey").SetFont("bold")
    currentGui.Add("Text", "x+60 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Modify").SetFont("bold")
    
    y := 50
    
    currentGui.Add("Text", "x20 y" y " w600 h2 0x10")
    y += 20
    currentGui.Add("Text", "x20 y" y " w200 c" (currentTheme = "dark" ? "White" : "Black"), "Applications").SetFont("bold s10")
    y += 30
    
    AddHotkeyRow("OpenExplorer", y, currentGui)
    y += 40
    AddHotkeyRow("OpenPowerShell", y, currentGui)
    y += 40
    AddHotkeyRow("OpenBrowser", y, currentGui)
    y += 40
    AddHotkeyRow("OpenVSCode", y, currentGui)
    y += 40
    AddHotkeyRow("OpenCalculator", y, currentGui)
    y += 40
    AddHotkeyRow("OpenSpotify", y, currentGui)
    
    y += 50
    currentGui.Add("Text", "x20 y" y " w600 h2 0x10")
    y += 20
    currentGui.Add("Text", "x20 y" y " w200 c" (currentTheme = "dark" ? "White" : "Black"), "System Tools").SetFont("bold s10")
    y += 30
    
    AddHotkeyRow("ToggleTaskbar", y, currentGui)
    y += 40
    AddHotkeyRow("ToggleDesktopIcons", y, currentGui)
    
    y += 50
    currentGui.Add("Text", "x20 y" y " w600 h2 0x10")
    y += 20
    currentGui.Add("Text", "x20 y" y " w200 c" (currentTheme = "dark" ? "White" : "Black"), "Sound Controls").SetFont("bold s10")
    y += 30
    
    AddHotkeyRow("VolumeUp", y, currentGui)
    y += 40
    AddHotkeyRow("VolumeDown", y, currentGui)
    y += 40
    AddHotkeyRow("ToggleMute", y, currentGui)
    y += 40
    AddHotkeyRow("ToggleMic", y, currentGui)
    
    y += 60
    currentGui.Add("Text", "x20 y" y " w600 h2 0x10")
    y += 20
    resetBtn := currentGui.Add("Button", "x500 y" y " w100 h30", "Reset All")
    resetBtn.SetFont("c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"))
    resetBtn.Opt("+Background" (currentTheme = "dark" ? "333333" : "DDDDDD"))
    resetBtn.OnEvent("Click", (*) => ResetAllHotkeys(currentGui))
    
    CloseKeybindsGUI(*) {
        global isMenuOpen
        isMenuOpen := false
        currentGui.Destroy()
    }
    
    currentGui.OnEvent("Close", CloseKeybindsGUI)
    
    currentGui.Show("w650 h790")
}

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
    global currentSetHotkeyGui, currentSetHotkeyAction
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
            global isMenuOpen := false
            ShowKeybindsGUI()
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
            global isMenuOpen := false
            ShowKeybindsGUI()
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
        Send("{Alt up}{Ctrl up}{Shift up}")
        Hotkey "Up", "Off"
        Hotkey "Down", "Off"
        Hotkey "Enter", "Off"
        Hotkey "Escape", "Off"
        
        Hotkey "Up", NavigateUp, "On"
        Hotkey "Down", NavigateDown, "On"
        Hotkey "Enter", ExecuteSelected, "On"
        Hotkey "Escape", (*) => CleanupEditorGui(), "On"
        return
    }
    
    CleanupEditorGui()
    
    isEditorMenuOpen := true
    currentSelection := 1
    
    Send("{Alt up}{Ctrl up}{Shift up}")
    
    hasVSCode := FindInPath("code.cmd") || FindInPath("code")
    hasVSCodium := FindInPath("codium.cmd") || FindInPath("codium")
    
    if (hasVSCode && hasVSCodium) {
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
        
        Hotkey "Up", NavigateUp, "On"
        Hotkey "Down", NavigateDown, "On"
        Hotkey "Enter", ExecuteSelected, "On"
        Hotkey "Escape", (*) => CleanupEditorGui(), "On"
        
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
        ShowKeybindsGUI()
    } else {
        gui.Show()
    }
}

CheckForUpdates() {
    global currentVersion, versionCheckUrl, githubReleasesUrl, isCheckingForUpdates
    
    if (isCheckingForUpdates)
        return
    
    isCheckingForUpdates := true
    
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
                    "Current version: " currentVersion "`n"
                    "Latest version: " latestVersion "`n`n"
                    "Would you like to visit the download page?",
                    "Update Available",
                    "YesNo 0x40"
                )
                if (result = "Yes") {
                    Run(githubReleasesUrl)
                }
            }
        } else {
            ShowNotification("❌ Failed to check for updates (Status: " whr.Status ")")
        }
    } catch Error as err {
        ShowNotification("❌ Failed to check for updates: " err.Message)
    }
    
    isCheckingForUpdates := false
}


ShowWelcomeGUI(*) {
    global welcomeGui
    
    if (IsObject(welcomeGui)) {
        welcomeGui.Destroy()
    }
    
    welcomeGui := Gui(, "WinMacros: Welcome")
    welcomeGui.SetFont("s10", "Segoe UI")
    welcomeGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "F0F0F0"
    
    if (currentTheme = "dark") {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", welcomeGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
    }
    
    welcomeGui.SetFont("s16 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x0 y20 w500 Center", "Welcome to WinMacros!")
    
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x0 y50 w500 Center", "Custom Windows macros for faster tasks and easy system control.")
    
    welcomeGui.SetFont("s10 bold", "Segoe UI")
    linkColor := currentTheme = "dark" ? "cWhite" : "cBlue"
    githubLink := welcomeGui.Add("Link", "x200 y+10 w500 Center -TabStop " linkColor, 
        '<a href="https://winmacros.netlify.app/">• Website</a> | <a href="https://github.com/fr0st-iwnl/WinMacros">• GitHub</a>')
    
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x20 y100 w460", "Helpful Shortcuts:")
    
    welcomeGui.Add("Text", "x20 y120 w460 h2 0x10")
    
    welcomeGui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x20 y130", "• Ctrl + Alt + K")
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x115 y130", "- Open Hotkey Settings")
    
    welcomeGui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x20 y+12", "• Ctrl + Alt + L")
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x115 y160", "- Open Launcher Settings")

    welcomeGui.SetFont("s10 bold c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x20 y+12", "• Alt + Backspace")
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x130 y190", "- Open Power Menu")
    
    welcomeGui.Add("Text", "x20 y+10 w460 h2 0x10")
    
    welcomeGui.SetFont("s10 norm c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    themeToggle := welcomeGui.Add("Checkbox", "x20 y+20 vThemeToggle", "Dark Theme")
    themeToggle.Value := currentTheme = "dark"
    themeToggle.OnEvent("Click", (*) => ToggleTheme(themeToggle.Value))
    
    startupPath := A_Startup "\WinMacros.lnk"
    isStartupEnabled := FileExist(startupPath) ? true : false
    startupChk := welcomeGui.Add("Checkbox", "x20 y+10 w200 vRunOnStartup c" (currentTheme = "dark" ? "White" : "Black"), "Run Script on Startup")
    startupChk.Value := isStartupEnabled
    startupChk.OnEvent("Click", ToggleWindowsStartup)
    
    showAtStartup := IniRead(settingsFile, "Settings", "ShowWelcome", "1")
    chk := welcomeGui.Add("Checkbox", "x280 y250 vShowAtStartup c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Show this message at startup")
    chk.Value := showAtStartup
    chk.OnEvent("Click", UpdateStartupState)
    
    welcomeGui.SetFont("s10", "Segoe UI")
    okBtn := welcomeGui.Add("Button", "x370 y300 w100", "OK")
    if (currentTheme = "dark") {
        okBtn.SetFont("c0xFFFFFF")
        okBtn.Opt("+Background333333")
    } else {
        okBtn.SetFont("c0x000000")
        okBtn.Opt("+BackgroundDDDDDD")
    }
    okBtn.OnEvent("Click", (*) => CloseWelcome(welcomeGui, startupChk.Value))
    
    welcomeGui.SetFont("s9 c" (currentTheme = "dark" ? "0xFFFFFF" : "0x000000"), "Segoe UI")
    welcomeGui.Add("Text", "x20 y305", "[#] Version: " currentVersion)
    
    welcomeGui.Show("w500 h350")
}

CloseWelcome(gui, showAgain) {
    IniWrite(showAgain, settingsFile, "Settings", "ShowWelcome")
    gui.Destroy()
}

if (IniRead(settingsFile, "Settings", "ShowWelcome", "1") = "1") {
    SetTimer(ShowWelcomeGUI, -100)
}

SetTimer(CheckForUpdates, -1000)

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
    global currentTheme, welcomeGui, currentSetHotkeyGui, currentSetHotkeyAction, isMenuOpen, powerGui, myGui
    currentTheme := isDark ? "dark" : "light"
    IniWrite(currentTheme, settingsFile, "Settings", "Theme")
    
    if (isDark)
        A_TrayMenu.Check("Dark Theme")
    else
        A_TrayMenu.Uncheck("Dark Theme")
    
    if (IsObject(welcomeGui)) {
        try {
            showAtStartup := welcomeGui["ShowAtStartup"].Value
            welcomeGui.Destroy()
            ShowWelcomeGUI()
            welcomeGui["ShowAtStartup"].Value := showAtStartup
            welcomeGui["ThemeToggle"].Value := isDark
        }
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
    
    if (IsObject(myGui) && WinExist("Launcher Settings")) {
        myGui.Destroy()
        ShowLauncherGUI()
    }
    
    if (isMenuOpen && WinExist("Hotkey Settings")) {
        isMenuOpen := false
        ShowKeybindsGUI()
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

^!l::ShowLauncherGUI()

LoadLaunchers()

ShowLauncherGUI() {
    global myGui
    
    if IsObject(myGui)
        myGui.Destroy()
    
    myGui := Gui(, "Launcher Settings")
    myGui.SetFont("s10", "Segoe UI")
    myGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "FFFFFF"
    
    if (currentTheme = "dark") {
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", myGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
        DllCall("uxtheme\SetWindowTheme", "Ptr", myGui.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    }
    
    myGui.Add("Text", "x10 y420 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Name:")
    nameInput := myGui.Add("Edit", "x110 y420 w200 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    
    myGui.Add("Text", "x10 y450 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Hotkey:")
    hotkeyInput := myGui.Add("Hotkey", "x110 y450 w200")
    helpText := myGui.Add("Link", "x+10 y453 w20 h20 -TabStop", '<a href="#">?</a>')
    helpText.SetFont("bold s10 underline c" (currentTheme = "dark" ? "0x98FB98" : "Blue"), "Segoe UI")
    helpText.OnEvent("Click", ShowHotkeyTooltip)
    
    ShowHotkeyTooltip(*) {
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
    
    myGui.Add("Text", "x10 y480 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Program Path:")
    pathInput := myGui.Add("Edit", "x110 y480 w500 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), "")
    browseBtn := myGui.Add("Button", "x620 y478 w100", "Browse")
    
    lv := myGui.Add("ListView", "x10 y10 w710 h400 Grid -Multi NoSortHdr +LV0x10000 +HScroll +VScroll", ["Name", "Hotkey", "Path"])
    
    if (currentTheme = "dark") {
        lv.Opt("+Background333333 cWhite")
        DllCall("uxtheme\SetWindowTheme", "Ptr", lv.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    }
    
    totalWidth := 710
    nameWidth := 150
    hotkeyWidth := 150
    pathWidth := totalWidth - nameWidth - hotkeyWidth - 4
    
    lv.ModifyCol(1, nameWidth)
    lv.ModifyCol(2, hotkeyWidth)
    lv.ModifyCol(3, pathWidth)
    
    addBtn := myGui.Add("Button", "x10 y520 w100", "Add")
    editBtn := myGui.Add("Button", "x120 y520 w100", "Edit")
    deleteBtn := myGui.Add("Button", "x230 y520 w100", "Delete")
    
    if (currentTheme = "dark") {
        browseBtn.Opt("+Background333333 cWhite")
        addBtn.Opt("+Background333333 cWhite")
        editBtn.Opt("+Background333333 cWhite")
        deleteBtn.Opt("+Background333333 cWhite")
    }
    
    LoadLaunchersToListView(lv)
    
    browseBtn.OnEvent("Click", BrowseFile)
    addBtn.OnEvent("Click", AddLauncher)
    editBtn.OnEvent("Click", EditLauncher)
    deleteBtn.OnEvent("Click", DeleteLauncher)
    myGui.OnEvent("Close", (*) => myGui.Hide())
    
    myGui.Show("w730 h560")
    
    BrowseFile(*) {
        if (selected := FileSelect(3,, "Select Program", "Programs (*.exe)")) {
            pathInput.Value := selected
        }
    }
    
    AddLauncher(*) {
        if (nameInput.Value = "" || hotkeyInput.Value = "" || pathInput.Value = "") {
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
                    
                    if (existingPath = pathInput.Value) {
                        ShowNotification("❌ A launcher with this path already exists")
                        return
                    }
                    if (existingName = nameInput.Value) {
                        ShowNotification("❌ A launcher with this name already exists")
                        return
                    }
                }
            }
        }
        
        IniWrite(pathInput.Value, launcherIniPath, "Launchers", hotkeyInput.Value)
        IniWrite(nameInput.Value, launcherIniPath, "Names", hotkeyInput.Value)
        
        CreateLauncherHotkey(hotkeyInput.Value, pathInput.Value)
        
        LoadLaunchersToListView(lv)
        
        nameInput.Value := ""
        hotkeyInput.Value := ""
        pathInput.Value := ""
        
        ShowNotification("✅ Launcher added successfully")
    }
    
    EditLauncher(*) {
        static editGuiOpen := false
        
        if (editGuiOpen) {
            return
        }
        
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
        
        editGui := Gui("+Owner" myGui.Hwnd, "Edit Launcher")
        editGui.SetFont("s10", "Segoe UI")
        editGui.BackColor := currentTheme = "dark" ? "1A1A1A" : "FFFFFF"
        
        if (currentTheme = "dark") {
            DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", editGui.Hwnd, "Int", 20, "Int*", true, "Int", 4)
            DllCall("uxtheme\SetWindowTheme", "Ptr", editGui.Hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
        }
        
        editGui.Add("Text", "x10 y10 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Name:")
        editNameInput := editGui.Add("Edit", "x110 y10 w200 h23 " (currentTheme = "dark" ? "Background333333 cWhite" : "BackgroundF0F0F0"), oldName)
        
        editGui.Add("Text", "x10 y40 w100 c" (currentTheme = "dark" ? "White" : "Black"), "Hotkey:")
        if (oldHotkey = "CapsLock" || oldHotkey = "Pause" || oldHotkey = "ScrollLock") {
            editHotkeyInput := editGui.Add("Hotkey", "x110 y40 w200", oldHotkey)
            editHotkeyInput.Enabled := true
        } else {
            editHotkeyInput := editGui.Add("Hotkey", "x110 y40 w200", oldHotkey)
        }
        
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
            
            IniDelete(launcherIniPath, "Launchers", oldHotkey)
            IniDelete(launcherIniPath, "Names", oldHotkey)
            
            IniWrite(editPathInput.Value, launcherIniPath, "Launchers", editHotkeyInput.Value)
            IniWrite(editNameInput.Value, launcherIniPath, "Names", editHotkeyInput.Value)
            
            editGui.Destroy()
            editGuiOpen := false
            
            ShowNotification("✅ Launcher updated successfully")
            
            IniWrite(1, launcherIniPath, "Temp", "ReopenLauncher")
            Reload()
        }
        
        editBrowseBtn.OnEvent("Click", BrowsePath)
        saveBtn.OnEvent("Click", SaveChanges)
        cancelBtn.OnEvent("Click", (*) => (editGui.Destroy(), editGuiOpen := false))
        editGui.OnEvent("Close", (*) => (editGui.Destroy(), editGuiOpen := false))
        
        editGui.Show("w640 h150")
    }
    
    DeleteLauncher(*) {
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
        
        LoadLaunchersToListView(lv)
        
        ShowNotification("🗑️ Deleted launcher: " name)
        
        IniWrite(1, launcherIniPath, "Temp", "ReopenLauncher")
        Reload()
    }
}

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
    if FileExist(path) {
        try {
            name := IniRead(launcherIniPath, "Names", hotkeyStr, path)
            
            fn := (*) => (Run(path), ShowNotification("🚀 Launching " name))
            activeHotkeys[hotkeyStr] := fn
            
            Hotkey hotkeyStr, fn
        }
    }
}

if (IniRead(launcherIniPath, "Temp", "ReopenLauncher", 0) = 1) {
    IniDelete(launcherIniPath, "Temp", "ReopenLauncher")
    SetTimer ShowLauncherGUI, -100
}