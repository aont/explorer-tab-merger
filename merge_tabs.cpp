// merge_tabs.cpp - Merge Explorer tabs into the first window (ANSI, MinGW-w64 friendly)
// Build: g++ merge_tabs.cpp -std=c++17 -lole32 -loleaut32 -lshell32 -lshlwapi -luuid -luser32

#define _WIN32_IE 0x0700
#define _WIN32_DCOM

#include <windows.h>
#include <shlobj.h>
#include <shlguid.h>
#include <exdisp.h>
#include <shldisp.h>
#include <servprov.h>
#include <oleauto.h>
#include <shobjidl.h>

#include <vector>
#include <string>
#include <algorithm>
#include <iostream>
#include <iomanip>

static const UINT WM_COMMAND_ID_NEW_TAB = 0xA21B; // same as newtab.cpp (undocumented)

struct TabInfo {
    IWebBrowser2* browser; // holds one reference; caller must Release
    std::wstring navTarget;
    std::string displayUrl;
    HWND topLevel;
};

// Get parsing name from PIDL (ANSI)
// Get parsing name from PIDL (UTF-8)
static std::string PidlToParsingName(PCIDLIST_ABSOLUTE pidl, bool* isFS){
    // Try normal filesystem path first (ANSI)
    CHAR sz[MAX_PATH];
    if (SHGetPathFromIDListA(pidl, sz)){
        if (isFS) *isFS = true;
        return sz; // regular filesystem path
    }

    // For virtual folders etc.: use the wide-char SHGetNameFromIDList
    PWSTR pwsz = nullptr;
    if (SUCCEEDED(SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &pwsz)) && pwsz){
        int wlen = lstrlenW(pwsz);
        std::string s;
        if (wlen > 0){
            int need = WideCharToMultiByte(CP_UTF8, 0, pwsz, wlen, nullptr, 0, nullptr, nullptr);
            if (need > 0){
                s.resize(need);
                WideCharToMultiByte(CP_UTF8, 0, pwsz, wlen, s.data(), need, nullptr, nullptr);
            }
        }
        CoTaskMemFree(pwsz);
        if (isFS) *isFS = false;
        return s; // virtual folder etc. (UTF-8)
    }

    if (isFS) *isFS = false;
    return "";
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return std::wstring();
    int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    if (len <= 0) return std::wstring();
    std::wstring out(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), &out[0], len);
    return out;
}

static std::string WideToUtf8(const std::wstring& ws) {
    if (ws.empty()) return std::string();
    int len = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    if (len <= 0) return std::string();
    std::string out(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), &out[0], len, nullptr, nullptr);
    return out;
}

// --- Helpers for BSTR/ANSI ---
static BSTR WideToBSTR(const std::wstring& s) {
    if (s.empty()) return nullptr;
    BSTR b = SysAllocStringLen(s.data(), (UINT)s.size());
    return b;
}

static std::wstring BSTRtoWide(BSTR b) {
    if (!b) return std::wstring();
    UINT lenW = SysStringLen(b);
    if (lenW == 0) return std::wstring();
    return std::wstring(b, b + lenW);
}

static HRESULT NavigateBrowser(IWebBrowser2* wb, const std::wstring& url) {
    if (!wb) return E_POINTER;
    VARIANT vURL; VariantInit(&vURL);
    VARIANT vEmpty; VariantInit(&vEmpty);

    vURL.vt = VT_BSTR;
    vURL.bstrVal = WideToBSTR(url);
    if (!vURL.bstrVal) {
        VariantClear(&vURL);
        return E_OUTOFMEMORY;
    }

    HRESULT hr = wb->Navigate2(&vURL, &vEmpty, &vEmpty, &vEmpty, &vEmpty);
    VariantClear(&vURL);
    VariantClear(&vEmpty);
    return hr;
}

// --- Collect Explorer tabs ---
static bool CollectExplorerTabs(std::vector<TabInfo>& tabs, std::vector<HWND>& windowOrder) {
    tabs.clear();
    windowOrder.clear();

    IShellWindows* pSW = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pSW)))) {
        return false;
    }

    long count = 0;
    if (FAILED(pSW->get_Count(&count))) {
        pSW->Release();
        return false;
    }

    for (long i = 0; i < count; ++i) {
        VARIANT vIdx; VariantInit(&vIdx);
        vIdx.vt = VT_I4;
        vIdx.lVal = i;

        IDispatch* pDisp = nullptr;
        if (FAILED(pSW->Item(vIdx, &pDisp)) || !pDisp) {
            VariantClear(&vIdx);
            continue;
        }
        VariantClear(&vIdx);

        IWebBrowser2* pWB = nullptr;
        if (FAILED(pDisp->QueryInterface(IID_IWebBrowser2, (void**)&pWB)) || !pWB) {
            pDisp->Release();
            continue;
        }
        pDisp->Release();

        bool isExplorer = false;
        std::wstring navTarget;
        std::string displayUrl;
        IServiceProvider* sp = nullptr;
        if (SUCCEEDED(pWB->QueryInterface(IID_IServiceProvider, (void**)&sp)) && sp) {
            IShellBrowser* sb = nullptr;
            if (SUCCEEDED(sp->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&sb))) && sb) {
                isExplorer = true;
                IShellView* sv = nullptr;
                if (SUCCEEDED(sb->QueryActiveShellView(&sv)) && sv) {
                    IFolderView* fv = nullptr;
                    if (SUCCEEDED(sv->QueryInterface(IID_PPV_ARGS(&fv))) && fv) {
                        IPersistFolder2* pf2 = nullptr;
                        if (SUCCEEDED(fv->GetFolder(IID_PPV_ARGS(&pf2))) && pf2) {
                            PIDLIST_ABSOLUTE pidl = nullptr;
                            if (SUCCEEDED(pf2->GetCurFolder(&pidl)) && pidl) {
                                bool isFS = false;
                                std::string parsingName = PidlToParsingName(pidl, &isFS);
                                if (!parsingName.empty()) {
                                    navTarget = Utf8ToWide(parsingName);
                                    displayUrl = parsingName;
                                }
                                CoTaskMemFree(pidl);
                            }
                            pf2->Release();
                        }
                        fv->Release();
                    }
                    sv->Release();
                }
                sb->Release();
            }
            sp->Release();
        }
        if (!isExplorer) {
            pWB->Release();
            continue;
        }

        SHANDLE_PTR handle = 0;
        HWND topLevel = nullptr;
        if (SUCCEEDED(pWB->get_HWND(&handle))) {
            topLevel = (HWND)handle;
        }
        if (!topLevel) {
            pWB->Release();
            continue;
        }

        if (navTarget.empty()) {
            BSTR bUrl = nullptr;
            if (SUCCEEDED(pWB->get_LocationURL(&bUrl)) && bUrl) {
                navTarget = BSTRtoWide(bUrl);
                SysFreeString(bUrl);
                displayUrl = WideToUtf8(navTarget);
            }
        }

        if (displayUrl.empty()) {
            displayUrl = WideToUtf8(navTarget);
        }

        if (std::find(windowOrder.begin(), windowOrder.end(), topLevel) == windowOrder.end()) {
            windowOrder.push_back(topLevel);
        }

        std::cout << "[debug] Explorer tab found: top-level HWND=0x" << std::hex << std::setw(0)
                  << reinterpret_cast<uintptr_t>(topLevel)
                  << ", IWebBrowser2=" << pWB
                  << ", URL=" << displayUrl << std::dec << "\n";

        tabs.push_back({ pWB, navTarget, displayUrl, topLevel });
    }

    pSW->Release();
    return true;
}

// --- Find ShellTabWindowClass inside a top-level Explorer window ---
struct FindTabHostData {
    HWND target;
};

static BOOL CALLBACK EnumFindTabHost(HWND hwnd, LPARAM lParam) {
    auto* data = reinterpret_cast<FindTabHostData*>(lParam);
    if (!data) return FALSE;

    char cls[256] = {0};
    if (GetClassNameA(hwnd, cls, 255) && std::string(cls) == "ShellTabWindowClass") {
        data->target = hwnd;
        return FALSE; // stop enumeration
    }

    EnumChildWindows(hwnd, EnumFindTabHost, lParam);
    return data->target ? FALSE : TRUE;
}

static HWND FindShellTabHost(HWND topLevel) {
    FindTabHostData data{};
    EnumChildWindows(topLevel, EnumFindTabHost, reinterpret_cast<LPARAM>(&data));
    return data.target;
}

// --- Create new tab in the first window and navigate ---
static bool CreateTabAndNavigate(HWND firstWindow, HWND tabHost, const std::wstring& url, size_t& knownTabCount) {
    if (!firstWindow || !tabHost || url.empty()) return false;

    size_t baselineCount = knownTabCount;
    {
        std::vector<TabInfo> beforeTabs;
        std::vector<HWND> beforeWindows;
        if (CollectExplorerTabs(beforeTabs, beforeWindows)) {
            size_t currentCount = 0;
            for (auto& t : beforeTabs) {
                if (t.topLevel == firstWindow) {
                    ++currentCount;
                }
            }
            baselineCount = currentCount;
            knownTabCount = currentCount;
            std::cout << "[debug] Baseline tab count for first window: " << baselineCount << "\n";
        }
        for (auto& t : beforeTabs) {
            if (t.browser) t.browser->Release();
        }
    }

    std::cout << "[debug] Sending WM_COMMAND to create new tab in HWND=0x" << std::hex
              << reinterpret_cast<uintptr_t>(tabHost) << std::dec << "\n";
    SendMessageA(tabHost, WM_COMMAND, (WPARAM)WM_COMMAND_ID_NEW_TAB, 0);

    const DWORD timeoutMs = 8000;
    const DWORD retryMs = 300;
    DWORD waited = 0;

    while (waited <= timeoutMs) {
        std::vector<TabInfo> tabs;
        std::vector<HWND> windows;
        if (!CollectExplorerTabs(tabs, windows)) {
            Sleep(retryMs);
            waited += retryMs;
            continue;
        }

        std::vector<TabInfo*> firstWindowTabs;
        for (auto& t : tabs) {
            if (t.topLevel == firstWindow) {
                firstWindowTabs.push_back(&t);
            }
        }

        size_t currentCount = firstWindowTabs.size();
        if (currentCount > baselineCount && !firstWindowTabs.empty()) {
            TabInfo* newestTab = firstWindowTabs.back();
            IWebBrowser2* newBrowser = newestTab->browser;
            std::cout << "[debug] Identified new tab by count increase (" << baselineCount
                      << " -> " << currentCount << ") in HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(firstWindow) << std::dec << "\n";

            HRESULT navHr = NavigateBrowser(newBrowser, url);

            for (auto& t : tabs) {
                if (t.browser && t.browser != newBrowser) {
                    t.browser->Release();
                }
            }

            if (SUCCEEDED(navHr)) {
                knownTabCount = currentCount;
                baselineCount = currentCount;
                std::cout << "[debug] Navigation succeeded for new tab.\n";
                newBrowser->Release();
                return true;
            }

            newBrowser->Release();
            return false;
        }

        for (auto& t : tabs) {
            if (t.browser) {
                t.browser->Release();
            }
        }

        Sleep(retryMs);
        waited += retryMs;
    }

    return false;
}

int main() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        std::cerr << "CoInitializeEx failed: 0x" << std::hex << hr << "\n";
        return 1;
    }

    std::vector<TabInfo> tabs;
    std::vector<HWND> windowOrder;
    if (!CollectExplorerTabs(tabs, windowOrder)) {
        std::cerr << "Failed to enumerate Explorer tabs.\n";
        CoUninitialize();
        return 2;
    }

    if (windowOrder.empty()) {
        std::cout << "No Explorer windows detected.\n";
        for (auto& t : tabs) if (t.browser) t.browser->Release();
        CoUninitialize();
        return 0;
    }

    HWND firstWindow = windowOrder.front();
    size_t knownTabCount = 0;
    std::vector<std::wstring> urlsToMerge;
    std::vector<HWND> windowsToClose;

    for (auto& t : tabs) {
        if (t.topLevel == firstWindow) {
            ++knownTabCount;
            std::cout << "[debug] Known tab in first window on startup: HWND=0x" << std::hex
                      << reinterpret_cast<uintptr_t>(t.topLevel)
                      << ", IWebBrowser2=" << t.browser << std::dec << "\n";
        } else {
            if (!t.navTarget.empty()) {
                urlsToMerge.push_back(t.navTarget);
                std::cout << "[debug] Tab queued for merge: HWND=0x" << std::hex
                          << reinterpret_cast<uintptr_t>(t.topLevel)
                          << ", IWebBrowser2=" << t.browser << std::dec
                          << ", URL=" << t.displayUrl << "\n";
            }
            if (std::find(windowsToClose.begin(), windowsToClose.end(), t.topLevel) == windowsToClose.end()) {
                windowsToClose.push_back(t.topLevel);
            }
        }
    }

    for (auto& t : tabs) {
        if (t.browser) t.browser->Release();
    }

    if (urlsToMerge.empty()) {
        std::cout << "Nothing to merge.\n";
        CoUninitialize();
        return 0;
    }

    HWND tabHost = FindShellTabHost(firstWindow);
    if (!tabHost) {
        std::cerr << "Could not find ShellTabWindowClass in the first window.\n";
        CoUninitialize();
        return 3;
    }

    std::cout << "Merging " << urlsToMerge.size() << " tab(s) into the first window...\n";

    size_t successCount = 0;
    for (const auto& url : urlsToMerge) {
        if (CreateTabAndNavigate(firstWindow, tabHost, url, knownTabCount)) {
            ++successCount;
        } else {
            std::cerr << "[warn] Failed to create tab for: " << WideToUtf8(url) << "\n";
        }
    }

    for (HWND h : windowsToClose) {
        if (h && h != firstWindow) {
            PostMessageA(h, WM_CLOSE, 0, 0);
        }
    }

    std::cout << "Completed. " << successCount << " tab(s) moved.\n";

    CoUninitialize();
    return 0;
}
