# Explorer Tab Merger
Explorer Tab Merger is a Windows utility that collects every open `explorer.exe` window and replays them as tabs inside the first window, effectively giving you a single tabbed Explorer shell.

## How it works
The tool automates the legacy COM automation interfaces behind File Explorer to enumerate every tab (`IShellWindows`/`IWebBrowser2`) and track which top-level window owns each one. It then locates the hidden `ShellTabWindowClass` child window in the primary Explorer window and sends it the undocumented `WM_COMMAND` message used by the native “New tab” button. Each newly created tab receives the original location through `IWebBrowser2::Navigate2`, and the now-empty donor windows are closed.

## Implementations
Explorer Tab Merger ships with both a native C++ implementation and a Python port. Pick whichever fits best with your tooling and deployment needs.

## Open a folder in a new tab
Need to jump to a specific folder without losing your existing File Explorer window? Use the companion utilities below to create a new tab in the first open Explorer window; if none exists, the tools fall back to `ShellExecute` to launch the folder directly. Both variants accept forward slashes (`/`) or backslashes (`\`) in the folder path.

### C++ version (`open_folder_tab.cpp`)
1. Build the executable with a MinGW-w64 toolchain (or Visual C++ with equivalent libraries):
   ```bash
   g++ open_folder_tab.cpp -std=c++17 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32 -o open_folder_tab.exe
   ```
2. Run the resulting binary with the folder you want to open:
   ```bash
   open_folder_tab.exe "C:/path/to/folder"
   ```

### Python version (`open_folder_tab.py`)
1. Install the required dependency (pywin32) into your Python environment:
   ```bash
   pip install pywin32
   ```
2. Launch the script with Python 3 and the folder path:
   ```bash
   python open_folder_tab.py "C:/path/to/folder"
   ```

### C++ version (`merge_tabs.cpp`)
1. Build the executable with a MinGW-w64 toolchain (or Visual C++ with equivalent libraries):
   ```bash
   g++ merge_tabs.cpp -std=c++17 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32 -o merge_tabs.exe
   ```
2. Run the resulting binary from a Command Prompt or PowerShell session while multiple Explorer windows are open:
   ```bash
   merge_tabs.exe
   ```
3. The program will merge every additional Explorer window into the first one, then close the redundant top-level windows.

### Python version (`merge_tabs.py`)
1. Install the required dependency (pywin32) into your Python environment:
   ```bash
   pip install pywin32
   ```
2. Launch the script with Python 3 while your target Explorer windows are open:
   ```bash
   python merge_tabs.py
   ```
3. The script mirrors the native logic: it opens new tabs inside the first Explorer window, navigates them to the original locations, and closes the donor windows.
