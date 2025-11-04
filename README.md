# Explorer Tab Merger

Explorer Tab Merger is a Windows utility that collects every open `explorer.exe` window and replays them as tabs inside the first window, effectively giving you a single tabbed Explorer shell.

## How it works

The tool automates the legacy COM automation interfaces behind File Explorer to enumerate every tab (`IShellWindows`/`IWebBrowser2`) and track which top-level window owns each one. It then locates the hidden `ShellTabWindowClass` child window in the primary Explorer window and sends it the undocumented `WM_COMMAND` message used by the native “New tab” button. Each newly created tab receives the original location through `IWebBrowser2::Navigate2`, and the now-empty donor windows are closed.
