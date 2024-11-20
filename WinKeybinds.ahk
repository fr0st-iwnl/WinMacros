;=====================================[ Welcome to My Windows 11 Keybindings ]====================================
;
; Welcome to my custom Windows 11 keybindings script, made with AutoHotkey!
; This script lets you use hotkeys for common actions to make your system easier to control.
; It includes features such as:
;   - [Audio Management] : Volume Control + Audio and Microphone muting
;   - [System Control] : Quick access to system shutdown, restart, lock, and sleep
;   - [Shortcuts] : Launch File Explorer, Powershell and more with custom hotkeys
;
; Enjoy :)
;
;=====================================[ Let's Get Started :P ]====================================

currentActiveMenu := ""

;---------------------------------[ VARIABLES / MAIN CONFIGURATION ]---------------------------------
fileExplorerKey     := "!e        # Alt + E"
powershellKey       := "!t        # Alt + T"
taskManagerKey      := "^+Esc     # Ctrl + Shift + Esc"
toggleTaskbarKey    := "!g        # Alt + G"
openBrowserKey      := "!f        # Alt + F"
muteAudioKey        := "Pause     # Pause/Break"
muteMicKey          := "Insert    # Insert"
increaseVolumeKey   := "!+Up      # Alt + Shift + Up"
decreaseVolumeKey   := "!+Down    # Alt + Shift + Down"
systemControlKey    := "!Backspace # Alt + Backspace"

;---------------------------------[ HOTKEYS / MAIN CONFIGURATION ]---------------------------------
Hotkey, % GetHotkey(fileExplorerKey),     OpenFileExplorer
Hotkey, % GetHotkey(powershellKey),       OpenPowerShell
Hotkey, % GetHotkey(taskManagerKey),      OpenTaskManager
Hotkey, % GetHotkey(toggleTaskbarKey),    ToggleTaskbarVisibility
Hotkey, % GetHotkey(openBrowserKey),      OpenDefaultBrowser
Hotkey, % GetHotkey(muteAudioKey),        MuteUnmuteAudio
Hotkey, % GetHotkey(muteMicKey),          MuteUnmuteMic
Hotkey, % GetHotkey(increaseVolumeKey),   IncreaseVolume
Hotkey, % GetHotkey(decreaseVolumeKey),   DecreaseVolume
Hotkey, % GetHotkey(systemControlKey),    ShowPowerOptions

;---------------------------------[ HELP GUI ]---------------------------------
!?::
    ; Check if a menu is already open
    if (currentActiveMenu != "") {
        ToolTip, Please close the "%currentActiveMenu%" menu before opening something else.
        SetTimer, RemoveToolTip, 1000
        return
    }

    currentActiveMenu := "Keybinds Help"

    Gui, Destroy

    Gui, +Resize +MinSize +MaxSize  ; Allow resizing but enforce a minimum size

    ; Add text and groups for keybinds help
    Gui, Add, Text, x20 y10 w350 h20, --- Here are your keybinds for quick and easy system control! ---
    Gui, Add, Text, x20 y35 w350 h30, -- To customize or change keybinds, edit the configuration in the script. --

    Gui, Add, GroupBox, x20 y60 w340 h50,  % Chr(0x1F527) " System Control"
    Gui, Add, Text, x40 y80 w300 h20, % GetHotkeyDisplay(systemControlKey) ": Open Power Options"

    Gui, Add, GroupBox, x20 y120 w340 h130,  % Chr(0x2328) " Shortcuts"
    Gui, Add, Text, x40 y140 w300 h20, % GetHotkeyDisplay(fileExplorerKey) ": Open File Explorer"
    Gui, Add, Text, x40 y160 w300 h20, % GetHotkeyDisplay(openBrowserKey) ": Open Browser"
    Gui, Add, Text, x40 y180 w300 h20, % GetHotkeyDisplay(powershellKey) ": Open Powershell"
    Gui, Add, Text, x40 y200 w300 h20, % GetHotkeyDisplay(toggleTaskbarKey) ": Toggle Taskbar Visibility"
    Gui, Add, Text, x40 y220 w300 h20, % GetHotkeyDisplay(taskManagerKey) ": Open Task Manager"

    Gui, Add, GroupBox, x20 y260 w340 h110, % Chr(0x1F50A) " Audio Manager"
    Gui, Add, Text, x40 y280 w300 h20, % GetHotkeyDisplay(muteAudioKey) ": Mute/Unmute Audio"
    Gui, Add, Text, x40 y300 w300 h20, % GetHotkeyDisplay(muteMicKey) ": Mute/Unmute Mic"
    Gui, Add, Text, x40 y320 w300 h20, % GetHotkeyDisplay(increaseVolumeKey) ": Increase Volume"
    Gui, Add, Text, x40 y340 w300 h20, % GetHotkeyDisplay(decreaseVolumeKey) ": Decrease Volume"

    Gui, Show, w380 h400, Keybinds Help
return

;---------------------------------[ AUDIO MANAGEMENT ]---------------------------------

MuteUnmuteAudio:
    ; Toggle audio mute state
    if (audioMuted := !audioMuted) {
        SoundSet, +1, MASTER, mute, 2  ; Replace 2 with your audio device ID <--- IMPORTANT
        tooltipText := "Audio Off"
    } else {
        SoundSet, -1, MASTER, mute, 2  ; Replace 2 with your audio device ID <--- IMPORTANT
        tooltipText := "Audio On"
    }
    ToolTip, %tooltipText%
    SetTimer, RemoveToolTip, 1000
return

MuteUnmuteMic:
    ; Toggle mic mute state
    if (micMuted := !micMuted) {
        SoundSet, +1, MASTER, mute, 3  ; Replace 3 with your audio device ID <--- IMPORTANT
        tooltipText := "Mic Off"
    } else {
        SoundSet, -1, MASTER, mute, 3  ; Replace 3 with your audio device ID <--- IMPORTANT
        tooltipText := "Mic On"
    }
    ToolTip, %tooltipText%
    SetTimer, RemoveToolTip, 1000
return

IncreaseVolume:
    Send, {Volume_Up}
return

DecreaseVolume:
    Send, {Volume_Down}
return

;---------------------------------[ SHORTCUTS ]---------------------------------

OpenFileExplorer:
    Run, explorer.exe
return

OpenPowerShell:
    Run, powershell.exe -NoExit -Command "cd $env:USERPROFILE\Downloads"
return

OpenTaskManager:
    Run, taskmgr.exe
return

ToggleTaskbarVisibility:
    HideShowTaskbar(hide := !hide)
return

OpenDefaultBrowser:
    Run, % DefaultBrowser()
return

;---------------------------------[ SYSTEM CONTROL ]---------------------------------

ShowPowerOptions:
    ; If another menu is open, show a tooltip and exit
    if (currentActiveMenu != "") {
        ToolTip, Please close the "%currentActiveMenu%" menu before opening something else.
        SetTimer, RemoveToolTip, 1000
        return
    }
    currentActiveMenu := "Power Options"

    ; Destroy any existing GUI if it exists
    Gui, Destroy

    Gui Font, q5 s10, Arial Unicode MS
    Gui, Add, Text, x20 y20 w220 h30, Choose an action:
    Gui, Add, ListBox, vActionListBox w200 h120, Restart|Shutdown|Sleep|Logoff
    Gui, Add, Button, gSelectAction w200 h30, Select
    Gui, Show, , Power Options
return

~Enter::
    ; Handle Enter key press in Power Options menu
    if !WinActive("Power Options")
        return
    Gui, Submit, NoHide
    GoSub, SelectAction
return

SelectAction:
    Gui, Submit
    selectedAction := ActionListBox
    Gui, Destroy

    ; Perform the selected action
    if (selectedAction = "Restart")
        Shutdown, 2
    else if (selectedAction = "Shutdown")
        Shutdown, 1
    else if (selectedAction = "Sleep")
        DllCall("PowrProf\SetSuspendState", "int", 0, "int", 0, "int", 0)
    else if (selectedAction = "Logoff")
        Shutdown, 0
return

;---------------------------------[ MISC FUNCTIONS ]---------------------------------

;─────────────────────────────────────────[Function to hide or show the taskbar with auto-hide]─────────────────────────────────────────
; Script rewritten with improvements. Original inspiration from: 
; https://www.autohotkey.com/boards/viewtopic.php?t=113777
; Thanks :)
HideShowTaskbar(hide) {
    static ABM_SETSTATE := 0xA, ABS_AUTOHIDE := 0x1, ABS_ALWAYSONTOP := 0x2
    if (hide is not integer)
        hide := !hide
    VarSetCapacity(APPBARDATA, size := 2 * A_PtrSize + 2 * 4 + 16 + A_PtrSize, 0)
    NumPut(size, APPBARDATA)
    NumPut(WinExist("ahk_class Shell_TrayWnd"), APPBARDATA, A_PtrSize)
    NumPut(hide ? ABS_AUTOHIDE : ABS_ALWAYSONTOP, APPBARDATA, size - A_PtrSize)
    DllCall("Shell32\SHAppBarMessage", UInt, ABM_SETSTATE, Ptr, &APPBARDATA)
}

;─────────────────────────────────────────[Function to Open default browser]─────────────────────────────────────────
; Original inspiration from: 
; https://www.autohotkey.com/board/topic/67330-how-to-open-default-web-browser/
; Thanks :)
DefaultBrowser() {
    RegRead, BrowserKeyName, HKEY_CURRENT_USER, Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice, Progid
    RegRead, BrowserFullCommand, HKEY_CLASSES_ROOT, %BrowserKeyName%\shell\open\command
    StringGetPos, pos, BrowserFullCommand, ",,1
    pos := --pos
    StringMid, BrowserPathandEXE, BrowserFullCommand, 2, %pos%
    Return BrowserPathandEXE
}

RemoveToolTip:
    SetTimer, RemoveToolTip, Off
    ToolTip
return

GetHotkey(keybind) {
    StringSplit, result, keybind, #
    return Trim(result1)
}

GetHotkeyDisplay(keybind) {
    StringSplit, result, keybind, #
    return Trim(result2)
}

~Esc::
    if (currentActiveMenu != "") {
        Gui, Destroy
        currentActiveMenu := ""
    }
return

GuiClose:
    Gui, Destroy
    currentActiveMenu := ""
    return
