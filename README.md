# 🚀 \\\ WinMacros
**WinMacros** is a lightweight, customizable script for enhancing ***Windows*** productivity with macros and hotkeys. Built using **AutoHotkey**, it offers quick access to system controls, audio management, shortcuts, and more.

> [!NOTE]
> This script works only with AutoHotkey v1.1 (deprecated)

## 📦 Installation 

### 1. Download the repository
- Clone this repository or download the ZIP.
```
git clone https://github.com/fr0st-iwnl/WinMacros.git
cd WinMacros
```
### 2.  Download and install AutoHotkey
- Make sure you install this version [AutoHotkey v1.1 (deprecated)](https://autohotkey.com/).

### 3. Run the script
- Double-click the `.ahk` file to start the script.

### 4. **Optional: Compile**
- Compile the `.ahk` script into an executable for standalone usage.

## 🔑 Customize Keybindings

To update the hotkeys:

1. **Open the script in a text editor**  
   - Right-click on the `.ahk` file and select **Edit Script**.

2. **Locate the keybinding variables**  
   - In the script, you’ll find a section named `;---------------------------------[ VARIABLES ]---------------------------------`.  
   Here, keybinding variables are defined, such as:

     ```ahk
     fileExplorerKey     := "!e        # Alt + E"
     powershellKey       := "!t        # Alt + T"
     ```

3. **Update the hotkey assignment**  
   - Change the key combination for any variable to your preferred hotkey.  
     For example, to change the **File Explorer** shortcut from <kbd>Alt+E</kbd> to <kbd>Ctrl+Shift+E</kbd>, update the variable:
     
     ```ahk
     fileExplorerKey := "^+e       # Ctrl+Shift+E"
     ```

4. **Save and reload the script**  
   - After editing, save the file and reload the script by right-clicking the running AutoHotkey icon in the system tray and selecting **Reload Script**.

## Previews

<div align="left"> <table> <tr> <td align="center"><b>Keybinds Help</b></td> <td align="center"><b>FindAudio</b></td> <td align="center"><b>Power Menu</b></td> </tr> <tr> <td><img src="https://raw.githubusercontent.com/fr0st-iwnl/WinMacros/refs/heads/master/Assets/keybindshelp.png" alt="Keybinds Help" style="width:300px;"/></td> <td><img src="https://raw.githubusercontent.com/fr0st-iwnl/WinMacros/refs/heads/master/Assets/findaudiopreview.png" alt="FindAudio" style="width:300px;"/></td> <td><img src="https://raw.githubusercontent.com/fr0st-iwnl/WinMacros/refs/heads/master/Assets/powermenu.png" alt="Power Menu"/></td> </tr> </table> </div>



## 🎹 Keybindings

<div align="left">

| Keys | Action |
| :--- | :--- |
| <kbd>Alt+T</kbd> | Open Powershell |
| <kbd>Alt+E</kbd> | Open File Explorer |
| <kbd>Alt+F</kbd> | Open Default Browser |
| <kbd>Alt+G</kbd> | Toggle Taskbar Visibility |
| <kbd>Ctrl+Shift+ESC</kbd> | Open Task Manager |
| <kbd>Insert</kbd> | Mute microphone |
| <kbd>Pause/Break</kbd> | Mute volume  |
| <kbd>Alt+Shift+Up</kbd> | Increase Volume  |
| <kbd>Alt+Shift+Down</kbd> | Decrease Volume  |
| <kbd>Alt+Backspace</kbd> | Open Power Menu  |
| <kbd>Alt+Shift+?</kbd> | Open Keybinds Help Menu  |

</div>


## 🤝 Contributions 

Feel free to fork this repository and submit pull requests. Contributions to improve functionality, documentation, or adding features are welcome.
