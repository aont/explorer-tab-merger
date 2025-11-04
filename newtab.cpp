#include <windows.h>
#include <tlhelp32.h>
#include <string>
#include <vector>
#include <iostream>
#include <sstream>
#include <iomanip>

static const UINT WM_COMMAND_ID_NEW_TAB = 0xA21B; // 41499: undocumented (subject to change in the future)

struct TabWin {
    HWND hwnd;
    DWORD pid;
};

static bool IsExplorerTabHost(HWND h) {
    char cls[256] = {0};
    if (!GetClassNameA(h, cls, 256)) return false;
    return std::string(cls) == "ShellTabWindowClass";
}

static BOOL CALLBACK EnumChildProc(HWND hwnd, LPARAM lParam) {
    auto* tabs = reinterpret_cast<std::vector<TabWin>*>(lParam);
    if (IsExplorerTabHost(hwnd)) {
        DWORD pid = 0;
        GetWindowThreadProcessId(hwnd, &pid);
        tabs->push_back({ hwnd, pid });
    }
    // Deep search (grandchildren and beyond)
    EnumChildWindows(hwnd, EnumChildProc, lParam);
    return TRUE;
}

static BOOL CALLBACK EnumTopProc(HWND hwnd, LPARAM lParam) {
    EnumChildWindows(hwnd, EnumChildProc, lParam);
    return TRUE;
}

static std::vector<TabWin> GatherAllShellTabWindows() {
    std::vector<TabWin> result;
    EnumWindows(EnumTopProc, reinterpret_cast<LPARAM>(&result));
    return result;
}

static bool SendNewTab(HWND tabHwnd) {
    LRESULT r = SendMessageA(tabHwnd, WM_COMMAND, (WPARAM)WM_COMMAND_ID_NEW_TAB, 0);
    // Since r can sometimes always be 0, it's hard to strictly determine success. For now, don't treat it as a failure.
    return (r != 0) || (GetLastError() == 0);
}

static std::string HwndToString(HWND h) {
    std::ostringstream ss;
    ss << "0x" << std::hex << std::uppercase << (uintptr_t)h;
    return ss.str();
}

static std::string PidToExe(DWORD pid) {
    if (pid == 0) return "(unknown)";
    std::string out = "(unknown)";
    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE) return out;

    PROCESSENTRY32 pe{};
    pe.dwSize = sizeof(pe);
    if (Process32First(snap, &pe)) {
        do {
            if (pe.th32ProcessID == pid) {
                out = pe.szExeFile;
                break;
            }
        } while (Process32Next(snap, &pe));
    }
    CloseHandle(snap);
    return out;
}

static void PrintUsage(const char* exe) {
    std::cout
        << "Usage:\n"
        << "  " << exe << " --list           : Enumerate ShellTabWindowClass\n"
        << "  " << exe << " --newtab         : Open a new tab in all found ShellTabWindowClass windows\n"
        << "  " << exe << " --newtab --first : Send only to the first one found\n";
}

int main(int argc, char* argv[]) {
    bool doList = false;
    bool doNewTab = false;
    bool onlyFirst = false;

    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        if (a == "--list") doList = true;
        else if (a == "--newtab") doNewTab = true;
        else if (a == "--first") onlyFirst = true;
        else if (a == "-h" || a == "--help" || a == "/?") {
            PrintUsage(argv[0]);
            return 0;
        }
    }

    if (!doList && !doNewTab) {
        PrintUsage(argv[0]);
        return 0;
    }

    auto tabs = GatherAllShellTabWindows();

    if (doList) {
        if (tabs.empty()) {
            std::cout << "[i] No ShellTabWindowClass found.\n";
        } else {
            std::cout << "[i] Found " << tabs.size() << " ShellTabWindowClass windows\n";
            for (size_t i = 0; i < tabs.size(); ++i) {
                const auto& t = tabs[i];
                std::cout << "  [" << i << "] hwnd=" << HwndToString(t.hwnd)
                          << ", pid=" << t.pid
                          << " (" << PidToExe(t.pid) << ")\n";
            }
        }
    }

    if (doNewTab) {
        if (tabs.empty()) {
            std::cerr << "[!] No target found (ShellTabWindowClass not found).\n";
            return 2;
        }
        bool okAny = false;
        for (size_t i = 0; i < tabs.size(); ++i) {
            const auto& t = tabs[i];
            bool ok = SendNewTab(t.hwnd);
            std::cout << "[send] hwnd=" << HwndToString(t.hwnd)
                      << " -> WM_COMMAND 0x" << std::hex << WM_COMMAND_ID_NEW_TAB
                      << (ok ? "  [OK]\n" : "  [NG]\n");
            if (ok) okAny = true;
            if (onlyFirst) break;
        }
        return okAny ? 0 : 1;
    }

    return 0;
}
