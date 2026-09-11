# Explorer Tab Merger
Explorer Tab Merger is a Windows utility that collects every open `explorer.exe` window and replays them as tabs inside the first window, effectively giving you a single tabbed Explorer shell.

## How it works
The tool automates the legacy COM automation interfaces behind File Explorer to enumerate every tab (`IShellWindows`/`IWebBrowser2`) and track which top-level window owns each one. It then locates the hidden `ShellTabWindowClass` child window in the primary Explorer window and sends it the undocumented `WM_COMMAND` message used by the native “New tab” button. Each newly created tab receives the original location through `IWebBrowser2::Navigate2`, and the now-empty donor windows are closed.

## Build
The included makefiles build `merge_tabs.exe` and `open_folder_tab.exe` together. From a MinGW-w64 shell, run:

```bash
mingw32-make -f Makefile.mingw
```

From a Visual Studio Developer Command Prompt, run:

```bat
nmake /f Makefile.msvc
```

Append `clean` to either command to remove the generated executables (and, for MSVC, object files).

## Open a folder in a new tab
Need to jump to a specific folder without losing your existing File Explorer window? Use the companion utility below to create a new tab in the first open Explorer window; if none exists, it falls back to `ShellExecute` to launch the folder directly. It accepts forward slashes (`/`) or backslashes (`\`) in the folder path.

### Open a folder (`open_folder_tab.c`)
1. Build the executable with a MinGW-w64 toolchain (or Visual C with equivalent libraries):
   ```bash
   gcc open_folder_tab.c explorer_tabs.c -std=c11 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32 -o open_folder_tab.exe
   ```
2. Run the resulting binary with the folder you want to open:
   ```bash
   open_folder_tab.exe "C:/path/to/folder"
   ```

### Merge Explorer windows (`merge_tabs.c`)
1. Build the executable with a MinGW-w64 toolchain (or Visual C with equivalent libraries):
   ```bash
   gcc merge_tabs.c explorer_tabs.c -std=c11 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32 -o merge_tabs.exe
   ```
2. Run the resulting binary from a Command Prompt or PowerShell session while multiple Explorer windows are open:
   ```bash
   merge_tabs.exe
   ```
3. The program will merge every additional Explorer window into the first one, then close the redundant top-level windows.
